(async () => {
  const { useMultiFileAuthState } = await import("@whiskeysockets/baileys");
  const { state, saveCreds } = await useMultiFileAuthState("test_auth");
  console.log("State keys:", Object.keys(state));
  await saveCreds();
  console.log("Creds saved");
})();
