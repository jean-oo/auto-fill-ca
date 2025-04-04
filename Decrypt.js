// ==UserScript==
// @name         Secure Form Filler
// @namespace    http://yourdomain.com
// @version      1.0
// @description  Secure form filling solution
// @author       YourName
// @match        https://prson-srpel.apps.cic.gc.ca/en/*
// @grant        GM_xmlhttpRequest
// @connect      raw.githubusercontent.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/crypto-js/4.1.1/crypto-js.min.js
// ==/UserScript==

(async () => {
    const REPO_URL = 'https://github.com/jean-oo/auto-fill-ca.git';
  
    
    const fetchResource = url => new Promise(resolve => {
        GM_xmlhttpRequest({
            method: "GET",
            url: url,
            onload: res => resolve(res.responseText)
        });
    });

   // get encrypted code and key
    const [encryptedCode, encryptionKey] = await Promise.all([
        fetchResource(REPO_URL + 'encrypted.js'),
        fetchResource(REPO_URL + 'key.txt')
    ]);

    // AES decryption
    const decrypt = (ciphertext, key) => {
        const bytes = CryptoJS.AES.decrypt(ciphertext, key);
        return bytes.toString(CryptoJS.enc.Utf8);
    };

    // decrypt and execute the core script
    const decryptedCode = decrypt(encryptedCode, encryptionKey);
    const script = document.createElement('script');
    script.textContent = decryptedCode;
    document.head.appendChild(script);

    console.log('Secure Form Filler loaded');
})();
