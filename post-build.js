const fs = require('fs');
const path = require('path');

// Define the paths to the child folder and the parent folder
const childFolder = path.join(__dirname, 'docs/browser');
const parentFolder = path.join(__dirname, 'docs');

// Function to move all files from child folder to parent folder
async function moveAllFiles() {
    try {
        const files = await fs.promises.readdir(childFolder);

        for (const file of files) {
            const currentPath = path.join(childFolder, file);
            const destinationPath = path.join(parentFolder, file);

            await fs.promises.rename(currentPath, destinationPath);
            console.log(`${file} was moved to ${parentFolder}`);
        }
    } catch (err) {
        console.error('Error moving files:', err);
    }
}

// Run the function
moveAllFiles();