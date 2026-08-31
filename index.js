const path = require("path");
const QRCode = require("qrcode");
const { createClient } = require("@supabase/supabase-js");

require("dotenv").config();

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const ACCOUNT_ID = process.env.ACCOUNT_ID; // uuid, kamu tentukan sendiri di .env
const ACCOUNT_USER_ID = process.env.ACCOUNT_USER_ID; // fk ke profiles.id
const ACCOUNT_SESSION_NAME = process.env.ACCOUNT_SESSION_NAME || "default-session";

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

async function updateAccount(fields) {
  const { error } = await supabase
    .from("wa_accounts")
    .update({ ...fields, updated_at: new Date().toISOString() })
    .eq("id", ACCOUNT_ID);
  if (error) console.error("Gagal update wa_accounts:", error.message);
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

async function start() {
  const {
    default: makeWASocket,
    useMultiFileAuthState,
    DisconnectReason,
  } = await import("@whiskeysockets/baileys");

  // folder auth lokal, khusus akun ini. TIDAK pernah dikirim ke Supabase.
  const authFolder = path.join(__dirname, "auth_info_baileys", ACCOUNT_ID);
  const { state, saveCreds } = await useMultiFileAuthState(authFolder);
  
  const pino = require("pino");
  const sock = makeWASocket({ 
    auth: state,
    logger: pino({ level: "silent" }),
    browser: ["ServerAutomation", "Chrome", "1.0.0"]
  });
  sock.ev.on("creds.update", saveCreds);

  sock.ev.on("connection.update", async (update) => {
    const { connection, qr, lastDisconnect } = update;

    // --- INI BAGIAN INTINYA: QR dari Baileys -> base64 -> Supabase ---
    if (qr) {
      const qrBase64 = await QRCode.toDataURL(qr);
      console.log("QR baru, mengirim ke Supabase...");
      await updateAccount({ status: "qr", qr_code_base64: qrBase64 });
    }
    // ------------------------------------------------------------------

    if (connection === "open") {
      const phone = sock.user?.id?.split(":")[0]?.split("@")[0] ?? null;
      console.log("Terhubung ke WhatsApp. Nomor:", phone);
      await updateAccount({ status: "connected", qr_code_base64: null, phone });
    }

    if (connection === "close") {
      const loggedOut = lastDisconnect?.error?.output?.statusCode === DisconnectReason.loggedOut;
      console.log("Koneksi tertutup. loggedOut =", loggedOut);
      await updateAccount({ status: loggedOut ? "logged_out" : "disconnected", qr_code_base64: null });
      if (!loggedOut) start(); // reconnect otomatis
    }
  });
}

(async () => {
  await ensureAccountRow();
  await start();
})();