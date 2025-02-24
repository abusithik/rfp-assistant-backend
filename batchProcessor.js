// server/scripts/batchProcessor.js
const path = require('path');
const fs = require('fs').promises;

// Load environment variables from the correct path
const dotenv = require('dotenv');
const envPath = path.join(__dirname, '..', '.env');
console.log('Loading .env from:', envPath);
const result = dotenv.config({ path: envPath });

if (result.error) {
    console.error('Error loading .env file:', result.error);
} else {
    console.log('Environment variables loaded successfully');
    console.log('PINECONE_API_KEY exists:', !!process.env.PINECONE_API_KEY);
    console.log('PINECONE_ENVIRONMENT exists:', !!process.env.PINECONE_ENVIRONMENT);
    console.log('PINECONE_INDEX_NAME exists:', !!process.env.PINECONE_INDEX_NAME);
}

const { processExcelRFP } = require('../excelProcessor');

const RFP_DOCS_DIR = path.join(__dirname, '../rfp_documents');

async function processBatchRFPs() {
    try {
        // Create rfp_documents directory if it doesn't exist
        await fs.mkdir(RFP_DOCS_DIR, { recursive: true });

        // Read all files from the directory
        const files = await fs.readdir(RFP_DOCS_DIR);
        const excelFiles = files.filter(file => 
            file.endsWith('.xlsx') || file.endsWith('.xls')
        );

        console.log(`Found ${excelFiles.length} Excel files to process`);

        // Process each file
        for (const file of excelFiles) {
            try {
                console.log(`Processing ${file}...`);
                
                const filePath = path.join(RFP_DOCS_DIR, file);
                const fileBuffer = await fs.readFile(filePath);
                
                const metadata = {
                    rfpId: `RFP-${path.parse(file).name}`,
                    title: file,
                    uploadDate: new Date().toISOString(),
                    category: 'batch-uploaded'
                };

                const result = await processExcelRFP(fileBuffer, metadata);
                console.log(`Successfully processed ${file}:`, result);

                // Move processed files to a 'processed' directory
                const processedDir = path.join(RFP_DOCS_DIR, 'processed');
                await fs.mkdir(processedDir, { recursive: true });
                await fs.rename(filePath, path.join(processedDir, file));

            } catch (error) {
                console.error(`Error processing ${file}:`, error);
                // Continue with next file even if one fails
            }
        }

        console.log('Batch processing completed');
    } catch (error) {
        console.error('Batch processing error:', error);
    }
}

// Run the batch processor
processBatchRFPs().then(() => {
    console.log('Script execution completed');
    process.exit(0);
}).catch(error => {
    console.error('Script failed:', error);
    process.exit(1);
});