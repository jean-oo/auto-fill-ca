const crypto = require('crypto');
const path = require('path');
const fs = require('fs');

function encrypt(text, key) {
    const cipher = crypto.createCipher('aes-256-cbc', key); //create an encryption object
    let encrypted = cipher.update(text, 'utf8', 'hex'); // encrypt the text, in hex format
    encrypted += cipher.final('hex');
    return encrypted;
}

const publicDir = path.join(__dirname, 'public');
if (!fs.existsSync(publicDir)) {
    fs.mkdirSync(publicDir);
}

try {
    const code = fs.readFileSync(
        path.join(__dirname, 'src/core-script.js'),
        { encoding: 'utf8' }
    );
    
const coreScript = fs.readFileSync(
    path.join(__dirname, 'core-script.js'),
    { encoding: 'utf8' }
);
const encryptionKey = crypto.randomBytes(32);  // Change this key
const encryptedCode = encrypt(cireScript, encryptionKey);
const iv = crypto.randomBytes(16); // apply to CBC mode, when the input is the same, will generate different encryption results to enhance security 

fs.writeFileSync(
        path.join(publicDir, 'encrypted.js'),
        `${iv.toString('hex')}:${encrypted}`
    );
    
    fs.writeFileSync(
        path.join(publicDir, 'key.txt'),
        key.toString('hex')
    );
console.log("Encrypted Code:", encryptedCode);
