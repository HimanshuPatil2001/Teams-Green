// Runs at document_start in the page world (MAIN) on teams.cloud.microsoft.
// Patches crypto.subtle.importKey to capture the MSAL AES-GCM cache encryption
// key as it's imported, so the popup can decrypt localStorage cache entries.
(function () {
  const _importKey = crypto.subtle.importKey.bind(crypto.subtle);
  window.__sgtCryptoKeys = [];

  crypto.subtle.importKey = async function (format, keyData, algorithm, extractable, usages) {
    const algName = algorithm?.name ?? algorithm;
    // Force extractable so we can export it later for cache decryption
    const forceExtractable = algName === 'AES-GCM' ? true : extractable;
    const key = await _importKey(format, keyData, algorithm, forceExtractable, usages);
    if (algName === 'AES-GCM') {
      window.__sgtCryptoKeys.push(key);
    }
    return key;
  };
})();
