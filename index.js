const path = require("path");
const QRCode = require("qrcode");
const { createClient } = require("@supabase/supabase-js");
const fs = require("fs");
const net = require("net");
const { WebSocketServer } = require("ws");

require("dotenv").config();

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const ACCOUNT_ID = process.env.ACCOUNT_ID; // uuid, kamu tentukan sendiri di .env
const ACCOUNT_USER_ID = process.env.ACCOUNT_USER_ID; // fk ke profiles.id
const ACCOUNT_SESSION_NAME = process.env.ACCOUNT_SESSION_NAME || "default-session";
const PREFERRED_PORT = Number(process.env.WHATSAPP_PORT || 3001);
const QR_TIMEOUT_MS = Number(process.env.WA_QR_TIMEOUT_MS || 60000);

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  console.error("SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY belum diisi di .env.");
  process.exit(1);
}
if (!ACCOUNT_ID) {
  console.error("ACCOUNT_ID belum diisi di .env. Isi dengan uuid bebas (kamu yang generate).");
  process.exit(1);
}
if (!ACCOUNT_USER_ID) {
  console.error("ACCOUNT_USER_ID belum diisi di .env. Isi dengan id (uuid) dari tabel profiles.");
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);
let sock;
let connectionState = "starting";
let assignedPort = PREFERRED_PORT;
let lastQrValue;

// ── Port detection ─────────────────────────────────────────────────────────
// Cek apakah port tersedia, kalau tidak minta OS assign port bebas
function resolvePort(preferred) {
  return new Promise((resolve) => {
    const probe = net.createServer();
    probe.listen(preferred, "127.0.0.1", () => {
      probe.close(() => resolve(preferred));
    });
    probe.on("error", () => {
      // Port sudah dipakai — minta OS assign port bebas
      const fallback = net.createServer();
      fallback.listen(0, "127.0.0.1", () => {
        const { port } = fallback.address();
        fallback.close(() => {
          console.warn(`Port ${preferred} sudah dipakai. Menggunakan port bebas: ${port}`);
          resolve(port);
        });
      });
    });
  });
}

// ── WebSocket server (port ditentukan setelah resolvePort) ─────────────────
let websocketServer;
const websocketReady = resolvePort(PREFERRED_PORT).then((port) => {
  assignedPort = port;
  websocketServer = new WebSocketServer({ host: "127.0.0.1", port });
  websocketServer.on("connection", connectionHandler);
  websocketServer.on("listening", () => {
    console.log(`Bridge WebSocket WhatsApp aktif di localhost:${port}.`);
  });
  websocketServer.on("error", (error) => {
    console.error("WebSocket error:", error.message);
    process.exit(1);
  });
  return new Promise((resolve, reject) => {
    websocketServer.once("listening", resolve);
    websocketServer.once("error", reject);
  });
});

function broadcast(payload) {
  if (!websocketServer) return;
  const message = JSON.stringify(payload);
  for (const client of websocketServer.clients) {
    if (client.readyState === 1) client.send(message);
  }
}

function connectionHandler(client) {
  client.send(JSON.stringify({
    type: "connection",
    status: connectionState,
    connected: connectionState === "open",
    message: connectionState === "open" ? "Sudah terhubung ke WhatsApp." : "WhatsApp belum terhubung.",
  }));

  client.on("message", async (rawMessage) => {
    try {
      const request = JSON.parse(rawMessage.toString());
      if (request.action === "health") {
        client.send(JSON.stringify({
          ok: true,
          connected: connectionState === "open",
          status: connectionState,
          message: connectionState === "open" ? "Sudah terhubung ke WhatsApp." : "WhatsApp belum terhubung.",
        }));
        return;
      }
      if (request.action !== "send" || connectionState !== "open") {
        client.send(JSON.stringify({ ok: false, error: "WhatsApp belum terhubung." }));
        return;
      }
      const { groupId, imagePath, caption } = request;
      if (!groupId) throw new Error("groupId wajib diisi.");
      if (imagePath) {
        if (!fs.existsSync(imagePath)) throw new Error("imagePath tidak ditemukan.");
        await sock.sendMessage(groupId, { image: fs.readFileSync(imagePath) });
      } else if (caption?.trim()) {
        await sock.sendMessage(groupId, { text: caption.trim() });
      } else {
        throw new Error("imagePath atau caption wajib diisi.");
      }
      client.send(JSON.stringify({ ok: true }));
    } catch (error) {
      client.send(JSON.stringify({ ok: false, error: error.message }));
    }
  });
}

// ── Supabase helpers ───────────────────────────────────────────────────────
async function updateAccount(fields) {
  const { error } = await supabase
    .from("wa_accounts")
    .update({ ...fields, updated_at: new Date().toISOString() })
    .eq("id", ACCOUNT_ID);
  if (error) console.error("Gagal update wa_accounts:", error.message);
}

async function clearWsPort() {
  try {
    const { error } = await supabase
      .from("wa_accounts")
      .update({ status: "disconnected", updated_at: new Date().toISOString() })
      .eq("id", ACCOUNT_ID);
    if (error) console.error("Gagal mengubah status akun saat shutdown:", error.message);
  } catch (_) {
    // best-effort, jangan sampai blokir proses exit
  }
}

/**
 * Pastikan row wa_accounts untuk ACCOUNT_ID ini ada. Kalau belum ada, insert lengkap
 * dengan semua atribut (id, user_id, session_name, status default). Kalau sudah ada,
 * tidak menimpa apa-apa (biar tidak reset status yang sedang berjalan).
 */
async function ensureAccountRow() {
  const { data: existing, error: selectError } = await supabase
    .from("wa_accounts")
    .select("id")
    .eq("id", ACCOUNT_ID)
    .maybeSingle();

  if (selectError) {
    console.error("Gagal cek row wa_accounts:", selectError.message);
    process.exit(1);
  }

  if (existing) {
    console.log(`Row wa_accounts (${ACCOUNT_ID}) sudah ada, dipakai langsung.`);
    return;
  }

  const { error: insertError } = await supabase.from("wa_accounts").insert({
    id: ACCOUNT_ID,
    user_id: ACCOUNT_USER_ID,
    session_name: ACCOUNT_SESSION_NAME,
    status: "disconnected",
    phone: null,
    qr_code_base64: null,
  });

  if (insertError) {
    console.error("Gagal insert row wa_accounts:", insertError.message);
    process.exit(1);
  }

  console.log(`Row wa_accounts baru dibuat: id=${ACCOUNT_ID}, session_name=${ACCOUNT_SESSION_NAME}`);
}

async function syncAuthFolderStatus(authFolder) {
  const sessionFiles = fs.readdirSync(authFolder, { withFileTypes: true })
    .filter((entry) => entry.isFile());
  if (sessionFiles.length === 0) return;

  await updateAccount({ status: "connected" });
  console.log(`Sesi auth ditemukan di ${authFolder}. Status akun diubah menjadi connected.`);
}

async function syncWhatsAppGroups() {
  const groups = await sock.groupFetchAllParticipating();
  const refreshedAt = new Date().toISOString();
  const groupRows = Object.values(groups).map((group) => ({
    group_jid: group.id,
    group_name: group.subject || group.id,
    is_active: true,
    wa_account_id: ACCOUNT_ID,
    last_refresh: refreshedAt,
  }));

  const { error: deactivateError } = await supabase
    .from("wa_groups")
    .update({ is_active: false, last_refresh: refreshedAt })
    .eq("wa_account_id", ACCOUNT_ID);
  if (deactivateError) throw deactivateError;

  if (groupRows.length === 0) {
    console.log("Tidak ada grup WhatsApp yang ditemukan untuk akun ini.");
    return;
  }

  const { error: upsertError } = await supabase
    .from("wa_groups")
    .upsert(groupRows, { onConflict: "wa_account_id,group_jid" });
  if (upsertError) throw upsertError;

  console.log(`${groupRows.length} grup WhatsApp berhasil disinkronkan.`);
}

// ── Graceful shutdown ──────────────────────────────────────────────────────
async function shutdown(signal) {
  console.log(`\nMenerima ${signal}. Membersihkan dan keluar...`);
  await clearWsPort();
  process.exit(0);
}

process.on("SIGINT", () => shutdown("SIGINT"));
process.on("SIGTERM", () => shutdown("SIGTERM"));
process.on("uncaughtException", async (err) => {
  console.error("Uncaught exception:", err);
  await clearWsPort();
  process.exit(1);
});

// ── Baileys socket ─────────────────────────────────────────────────────────
async function start() {
  const {
    default: makeWASocket,
    useMultiFileAuthState,
    DisconnectReason,
  } = await import("@whiskeysockets/baileys");

  // folder auth lokal, khusus akun ini. TIDAK pernah dikirim ke Supabase.
  const authFolder = path.join(__dirname, "auth_info_baileys", ACCOUNT_ID);
  fs.mkdirSync(authFolder, { recursive: true });
  console.log("Folder kredensial Baileys lokal:", authFolder);
  await syncAuthFolderStatus(authFolder);
  const { state, saveCreds } = await useMultiFileAuthState(authFolder);
  
  const pino = require("pino");
  sock = makeWASocket({ 
    auth: state,
    logger: pino({ level: "silent" }),
    browser: ["ServerAutomation", "Chrome", "1.0.0"],
    qrTimeout: QR_TIMEOUT_MS,
  });
  sock.ev.on("creds.update", async (credentials) => {
    await saveCreds(credentials);
    console.log("Kredensial Baileys disimpan di:", authFolder);
  });

  sock.ev.on("connection.update", async (update) => {
    const { connection, qr, lastDisconnect } = update;

    // --- INI BAGIAN INTINYA: QR dari Baileys -> base64 -> Supabase ---
    if (qr) {
      if (qr !== lastQrValue) {
        lastQrValue = qr;
        console.log("QR login tersedia. Buka web admin panel untuk scan QR dan membuat sesi auth_info_baileys.");
        const qrBase64 = await QRCode.toDataURL(qr);
        console.log("QR baru, mengirim ke Supabase...");
        await updateAccount({ status: "qr", qr_code_base64: qrBase64 });
        broadcast({ type: "qr", status: "qr", connected: false, message: "QR login WhatsApp tersedia." });
      }
    }
    // ------------------------------------------------------------------

    if (connection === "open") {
      connectionState = "open";
      const phone = sock.user?.id?.split(":")[0]?.split("@")[0] ?? null;
      await saveCreds();
      console.log("Sudah terhubung ke WhatsApp. Nomor:", phone);
      await updateAccount({ status: "connected", qr_code_base64: null, phone });
      try {
        await syncWhatsAppGroups();
      } catch (error) {
        console.error("Gagal menyinkronkan grup WhatsApp:", error.message);
      }
      broadcast({
        type: "connection",
        status: "connected",
        connected: true,
        phone,
        message: "Sudah terhubung ke WhatsApp.",
      });
    }

    if (connection === "close") {
      connectionState = "close";
      const loggedOut = lastDisconnect?.error?.output?.statusCode === DisconnectReason.loggedOut;
      console.log("Koneksi tertutup. loggedOut =", loggedOut);
      await updateAccount({ status: loggedOut ? "logged_out" : "disconnected", qr_code_base64: null });
      broadcast({
        type: "connection",
        status: loggedOut ? "logged_out" : "disconnected",
        connected: false,
        message: loggedOut ? "Sesi WhatsApp sudah logout." : "Koneksi WhatsApp terputus.",
      });
      if (!loggedOut) start(); // reconnect otomatis
    }
  });
}

(async () => {
  try {
    await websocketReady;
    await ensureAccountRow();
    await start();
  } catch (error) {
    if (error.code !== "EADDRINUSE") {
      console.error("Bridge gagal dimulai:", error.message);
      process.exit(1);
    }
  }
})();