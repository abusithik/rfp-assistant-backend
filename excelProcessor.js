// server/excelProcessor.js

// Check required environment variables
const requiredEnvVars = ['PINECONE_API_KEY', 'PINECONE_ENVIRONMENT', 'PINECONE_INDEX_NAME', 'OPENAI_API_KEY'];
const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingVars.length > 0) {
    throw new Error(`Missing required environment variables: ${missingVars.join(', ')}`);
}

// Environment variable logging
console.log('Environment Check:');
console.log('PINECONE_API_KEY length:', process.env.PINECONE_API_KEY?.length || 0);
console.log('PINECONE_ENVIRONMENT:', process.env.PINECONE_ENVIRONMENT);
console.log('PINECONE_INDEX_NAME:', process.env.PINECONE_INDEX_NAME);

const ExcelJS = require('exceljs');
const _ = require('lodash');
const { Pinecone } = require('@pinecone-database/pinecone');
const { OpenAI } = require('openai');
const fetch = require('node-fetch');
const https = require('https');
const crypto = require('crypto');

// Initialize OpenAI
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
});

// Initialize Pinecone with ONLY the required properties
let pineconeClient = null;
let pineconeIndex = null;
try {
    pineconeClient = new Pinecone({
        apiKey: process.env.PINECONE_API_KEY,
        environment: process.env.PINECONE_ENVIRONMENT
    });
    pineconeIndex = pineconeClient.index(process.env.PINECONE_INDEX_NAME);
    console.log('Initialized Pinecone client with SDK');
} catch (error) {
    console.error('Failed to initialize Pinecone client:', error.message);
    console.log('Will fall back to direct API access');
}

// Create a custom agent with relaxed SSL verification for direct API access
const agent = new https.Agent({
    rejectUnauthorized: false,
    keepAlive: true,
    timeout: 60000
});

// Direct API implementation for Pinecone
const pineconeApi = {
    query: async (embedding, topK = 5, filters = {}) => {
        try {
            const baseUrl = `https://${process.env.PINECONE_INDEX_NAME}.svc.${process.env.PINECONE_ENVIRONMENT}.pinecone.io`;
            console.log('Querying Pinecone at URL:', baseUrl);
            
            const response = await fetch(`${baseUrl}/query`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Api-Key': process.env.PINECONE_API_KEY
                },
                body: JSON.stringify({
                    vector: embedding,
                    topK,
                    includeMetadata: true,
                    filter: Object.keys(filters).length > 0 ? filters : undefined
                }),
                agent,
                timeout: 30000
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Pinecone API error (${response.status}): ${errorText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('Pinecone API query error:', error);
            throw error;
        }
    },
    
    upsert: async (vectors) => {
        try {
            const baseUrl = `https://${process.env.PINECONE_INDEX_NAME}.svc.${process.env.PINECONE_ENVIRONMENT}.pinecone.io`;
            console.log('Upserting to Pinecone at URL:', baseUrl);
            
            const response = await fetch(`${baseUrl}/vectors/upsert`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Api-Key': process.env.PINECONE_API_KEY
                },
                body: JSON.stringify({
                    vectors
                }),
                agent,
                timeout: 30000
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Pinecone API error (${response.status}): ${errorText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('Pinecone API upsert error:', error);
            throw error;
        }
    },
    
    fetch: async (ids) => {
        try {
            const baseUrl = `https://${process.env.PINECONE_INDEX_NAME}.svc.${process.env.PINECONE_ENVIRONMENT}.pinecone.io`;
            console.log('Fetching from Pinecone at URL:', baseUrl);
            
            const response = await fetch(`${baseUrl}/vectors/fetch`, {
                method: 'POST', // Changed from GET to POST
                headers: {
                    'Content-Type': 'application/json',
                    'Api-Key': process.env.PINECONE_API_KEY
                },
                body: JSON.stringify({
                    ids
                }),
                agent,
                timeout: 30000
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Pinecone API error (${response.status}): ${errorText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('Pinecone API fetch error:', error);
            throw error;
        }
    }
};

async function testPineconeConnectivity() {
    try {
        const baseUrl = `https://${process.env.PINECONE_INDEX_NAME}.svc.${process.env.PINECONE_ENVIRONMENT}.pinecone.io`;
        console.log('Testing connectivity to Pinecone at:', baseUrl);
        
        // Try with a more permissive HTTPS agent
        const httpsAgent = new https.Agent({
            rejectUnauthorized: false,
            keepAlive: true,
            timeout: 30000,
            maxSockets: 5
        });
        
        // Add a direct curl-like debug to see what's happening
        console.log(`curl -v -H "Api-Key: ${process.env.PINECONE_API_KEY.substring(0, 4)}..." ${baseUrl}/describe_index_stats`);
        
        const response = await fetch(`${baseUrl}/describe_index_stats`, {
            method: 'GET',
            headers: {
                'Api-Key': process.env.PINECONE_API_KEY,
                'Accept': 'application/json'
            },
            agent: httpsAgent,
            timeout: 30000
        });
        
        if (response.ok) {
            const data = await response.json();
            console.log('Pinecone connectivity test successful!');
            console.log('Index stats:', JSON.stringify(data).substring(0, 200) + '...');
            return true;
        } else {
            console.error('Pinecone connectivity test failed with status:', response.status);
            const errorText = await response.text();
            console.error('Error details:', errorText);
            return false;
        }
    } catch (error) {
        console.error('Pinecone connectivity test failed with error:', error.message);
        if (error.code) {
            console.error('Error code:', error.code);
        }
        return false;
    }
}
// Run the test on startup
testPineconeConnectivity().then(isConnected => {
    if (!isConnected) {
        console.warn('WARNING: Unable to connect to Pinecone. The application may not work correctly.');
    }
});

// Add retry logic for operations
const withRetry = async (operation, maxRetries = 3, delay = 1000) => {
    for (let i = 0; i < maxRetries; i++) {
        try {
            return await operation();
        } catch (error) {
            if (i === maxRetries - 1) throw error;
            
            console.log(`Operation failed, attempt ${i + 1}/${maxRetries}. Retrying in ${delay}ms...`);
            console.error('Error:', error.message);
            
            await new Promise(resolve => setTimeout(resolve, delay));
            delay *= 2; // Exponential backoff
        }
    }
};

// Get embeddings from OpenAI
async function getEmbedding(text) {
    const response = await openai.embeddings.create({
        model: "text-embedding-ada-002",
        input: text,
    });
    return response.data[0].embedding;
}

// Generate a stable, unique ID for a piece of content
function generateStableId(metadata, item) {
    const content = `${metadata.rfpId}-${item.sheetName}-${item.category}-${item.text.slice(0, 50)}`;
    return crypto.createHash('md5').update(content).digest('hex');
}

async function processExcelRFP(buffer, metadata) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const processedData = [];
        let totalRows = 0;
        let processedRows = 0;
        let skippedRows = 0;
        let errorRows = 0;

        // Process each worksheet
        for (const worksheet of workbook.worksheets) {
            console.log(`Processing worksheet: ${worksheet.name}`);
            const sheetName = worksheet.name;
            const jsonData = [];

            // Get headers and clean them
            const headers = worksheet.getRow(1).values
                .slice(1) // Skip first empty cell
                .map(header => header ? header.trim() : '');

            // Count total rows
            totalRows += worksheet.rowCount - 1; // Exclude header row

            // Process each row
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { // Skip header row
                    const rowData = {};
                    row.eachCell((cell, colNumber) => {
                        const header = headers[colNumber - 1];
                        if (header) { // Only process cells with valid headers
                            rowData[header] = cell.text.trim();
                        }
                    });
                    if (Object.keys(rowData).length > 0) {
                        jsonData.push(rowData);
                    }
                }
            });

            // Group data by category if available
            const groupedData = _.groupBy(jsonData, 'Category');

            // Process each category
            for (const [category, items] of Object.entries(groupedData)) {
                for (const item of items) {
                    const combinedText = Object.entries(item)
                        .filter(([key, value]) => value && typeof value === 'string')
                        .map(([key, value]) => `${key}: ${value}`)
                        .join('\n');

                    if (combinedText.trim()) {
                        processedData.push({
                            category: category || 'uncategorized',
                            sheetName,
                            text: combinedText,
                            originalData: item
                        });
                    }
                }
            }
        }

        // Store in Pinecone with error handling, duplicate prevention, and batching
        console.log(`Total items to process: ${processedData.length}`);
        
        // Process in batches of 10
        const BATCH_SIZE = 10;
        const batches = _.chunk(processedData, BATCH_SIZE);
        
        for (const batch of batches) {
            const batchOperations = [];
            
            for (const item of batch) {
                try {
                    const vectorId = generateStableId(metadata, item);
                    
                    // Try to fetch existing vector with error handling - Use direct API if SDK failed
                    let existingVector;
                    try {
                        if (pineconeIndex) {
                            const fetchResponse = await withRetry(() => pineconeIndex.fetch([vectorId]));
                            existingVector = fetchResponse.vectors || {};
                        } else {
                            const fetchResponse = await withRetry(() => pineconeApi.fetch([vectorId]));
                            existingVector = fetchResponse.vectors || {};
                        }
                    } catch (fetchError) {
                        console.log(`Fetch check failed for ${vectorId}, proceeding with upsert`);
                        existingVector = {};
                    }
                    
                    if (!existingVector[vectorId]) {
                        const embedding = await withRetry(() => getEmbedding(item.text));
                        
                        batchOperations.push({
                            id: vectorId,
                            values: embedding,
                            metadata: {
                                ...metadata,
                                category: item.category,
                                sheetName: item.sheetName,
                                text: item.text,
                                originalData: JSON.stringify(item.originalData)
                            }
                        });
                        processedRows++;
                        console.log(`Prepared item ${processedRows}/${processedData.length} from ${item.sheetName}`);
                    } else {
                        skippedRows++;
                        console.log(`Skipping duplicate entry (${skippedRows} skipped so far)`);
                    }
                } catch (error) {
                    errorRows++;
                    console.error(`Error preparing item (${errorRows} errors so far):`, error.message);
                    continue;
                }
            }
            
            // Upload the batch with retry logic - Use direct API if SDK failed
            if (batchOperations.length > 0) {
                try {
                    if (pineconeIndex) {
                        await withRetry(() => pineconeIndex.upsert(batchOperations));
                    } else {
                        await withRetry(() => pineconeApi.upsert(batchOperations));
                    }
                    console.log(`Successfully uploaded batch of ${batchOperations.length} items`);
                } catch (batchError) {
                    console.error(`Error uploading batch: ${batchError.message}`);
                    errorRows += batchOperations.length;
                }
            }
        }

        return {
            success: true,
            stats: {
                totalItems: processedData.length,
                processed: processedRows,
                skipped: skippedRows,
                errors: errorRows
            },
            sheets: workbook.worksheets.map(sheet => sheet.name)
        };
    } catch (error) {
        console.error('Error processing Excel file:', error);
        throw error;
    }
}

async function queryRFPData(question, filters = {}) {
    try {
        // Get embedding from OpenAI
        const queryEmbedding = await withRetry(() => getEmbedding(question));
        console.log('Generated embedding with length:', queryEmbedding.length);

        // Prepare filter conditions
        let filterConditions = {};
        if (Object.keys(filters).length > 0) {
            if (filters.category) filterConditions.category = filters.category;
            if (filters.sheetName) filterConditions.sheetName = filters.sheetName;
        } else {
            filterConditions = {
                category: { $exists: true }
            };
        }

        // Try SDK first, fall back to direct API
        let queryResponse;
        try {
            if (pineconeIndex) {
                console.log('Querying Pinecone using SDK...');
                queryResponse = await withRetry(() => 
                    pineconeIndex.query({
                        vector: queryEmbedding,
                        topK: 5,
                        includeMetadata: true,
                        ...(Object.keys(filterConditions).length > 0 && { filter: filterConditions })
                    })
                );
            } else {
                throw new Error('SDK not initialized, using direct API');
            }
        } catch (sdkError) {
            console.log('SDK query failed, falling back to direct API:', sdkError.message);
            queryResponse = await withRetry(() => 
                pineconeApi.query(queryEmbedding, 5, filterConditions)
            );
        }
        
        console.log('Query response received:', !!queryResponse);

        // Handle potential empty responses
        if (!queryResponse.matches || queryResponse.matches.length === 0) {
            console.log('No matches found in Pinecone');
            
            // If this is a greeting or general question, respond appropriately
            if (question.toLowerCase().includes('hi') || 
                question.toLowerCase().includes('hello') || 
                question.toLowerCase().includes('hey')) {
                return {
                    answer: "Hello! I'm your RFP Assistant. I can help you find information in your RFP documents. How can I assist you today?",
                    sources: []
                };
            }
            
            return {
                answer: "I couldn't find any relevant information for your question in the RFP documents. Could you try rephrasing your question or ask about a different aspect of the RFP?",
                sources: []
            };
        }

        // If we have a valid response with matches
        const contexts = queryResponse.matches.map(match => ({
            text: match.metadata.text,
            originalData: JSON.parse(match.metadata.originalData),
            category: match.metadata.category,
            sheetName: match.metadata.sheetName,
            similarity: match.score
        }));

        // Generate response using OpenAI
        const completion = await withRetry(() =>
            openai.chat.completions.create({
                model: "gpt-4",
                messages: [
                    {
                        role: "system",
                        content: `You are an RFP assistant specialized in analyzing historical RFP data

For general queries and greetings:
- Respond in a friendly, professional manner
- Introduce yourself as the RFP Assistant
- Be precise with the answers unless asked you to explain.

For RFP-specific queries:
- Provide precise answers based on the provided context
- Include specific details from the data when relevant but not quote from where you are finding the information
- Highlight key information and requirements
- Always maintain a professional yet friendly tone
- If the question is not RFP-related, engage appropriately while gently guiding the conversation toward RFP topics`
                    },
                    {
                        role: "user",
                        content: `Context from RFP data:\n${contexts.map(c => 
                            `[Sheet: ${c.sheetName}, Category: ${c.category}]\n${c.text}`
                        ).join('\n\n')}\n\nQuestion: ${question}`
                    }
                ]
            })
        );

        return {
            answer: completion.choices[0].message.content,
            sources: contexts
        };
    } catch (error) {
        console.error('Error querying RFP data:', error);
        
        // Special handling for greeting messages even if there's an error
        if (question.toLowerCase().includes('hi') || 
            question.toLowerCase().includes('hello') || 
            question.toLowerCase().includes('hey')) {
            return {
                answer: "Hello! I'm your RFP Assistant. I can help you find information in your RFP documents. How can I assist you today?",
                sources: []
            };
        }
        
        // Provide a graceful fallback response
        return {
            answer: "I'm here to help with your RFP questions, but I'm having trouble connecting to my database right now. Could you please try again in a moment?",
            error: error.message,
            sources: []
        };
    }
}

module.exports = {
    processExcelRFP,
    queryRFPData
};
