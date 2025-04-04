const crypto = require('crypto');
const path = require('path');
const fs = require('fs');

function encrypt(text, key) {
    const cipher = crypto.createCipher('aes-256-cbc', key); //create an encryption object
    let encrypted = cipher.update(text, 'utf8', 'hex'); // encrypt the text, in hex format
    encrypted += cipher.final('hex');
    return encrypted;
}

const code = fs.readFileSync(
        path.join(__dirname, 'core-script.js'),
        { encoding: 'utf8' }
    );
const encryptionKey = crypto.randomBytes(32);  
const encryptedCode = encrypt(code, encryptionKey);
const iv = crypto.randomBytes(16); // apply to CBC mode, when the input is the same, will generate different encryption results to enhance security 

fs.writeFileSync(
        path.join(__dirname, 'encrypted.js'),
        `${iv.toString('hex')}:${encrypted}`
    );
    
    fs.writeFileSync(
        path.join(__dirname, 'key.txt'),
        key.toString('hex')
    );
console.log("Encrypted Code:", encryptedCode);
