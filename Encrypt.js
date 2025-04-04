const crypto = require('crypto');
const path = require('path');
const fs = require('fs');

function encrypt(text, key) {
    const cipher = crypto.createCipher('aes-256-cbc', key);
    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return encrypted;
}


const jsCode = fs.readFileSync(
    path.join(__dirname, 'core-script.js'),
    { encoding: 'utf8' }
);
const encryptionKey = crypto.randomBytes(32);  // Change this key
const encryptedCode = encrypt(jsCode, encryptionKey);

console.log("Encrypted Code:", encryptedCode);
