// server/excelProcessor.js

// Check required environment variables
const requiredEnvVars = ['PINECONE_API_KEY', 'PINECONE_ENVIRONMENT', 'PINECONE_INDEX_NAME', 'OPENAI_API_KEY'];
const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingVars.length > 0) {
    throw new Error(`Missing required environment variables: ${missingVars.join(', ')}`);
}

// Environment variable logging
console.log('Environment Check:');
console.log('PINECONE_API_KEY exists:', !!process.env.PINECONE_API_KEY);
console.log('PINECONE_ENVIRONMENT:', process.env.PINECONE_ENVIRONMENT);
console.log('PINECONE_INDEX_NAME:', process.env.PINECONE_INDEX_NAME);
console.log('PINECONE_HOST exists:', !!process.env.PINECONE_HOST);
console.log('OPENAI_API_KEY exists:', !!process.env.OPENAI_API_KEY);

const ExcelJS = require('exceljs');
const _ = require('lodash');
const { OpenAI } = require('openai');
const fetch = require('node-fetch');
const https = require('https');
const crypto = require('crypto');

// Initialize OpenAI
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
});

// Create a custom HTTPS agent with relaxed settings
const httpsAgent = new https.Agent({
    rejectUnauthorized: false,
    keepAlive: true,
    timeout: 60000
});

// Set up Pinecone base URL - prioritize direct host if available
const getPineconeBaseUrl = () => {
    if (process.env.PINECONE_HOST) {
        // If PINECONE_HOST includes the protocol, use it as is
        if (process.env.PINECONE_HOST.startsWith('http')) {
            return process.env.PINECONE_HOST;
        }
        // Otherwise, add the https:// prefix
        return `https://${process.env.PINECONE_HOST}`;
    }
    // Fall back to constructing the URL from environment and index name
    return `https://${process.env.PINECONE_INDEX_NAME}.svc.${process.env.PINECONE_ENVIRONMENT}.pinecone.io`;
};

// Direct API implementation for Pinecone
const pineconeApi = {
    query: async (embedding, topK = 5, filters = {}) => {
        try {
            const baseUrl = getPineconeBaseUrl();
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
                agent: httpsAgent,
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
            const baseUrl = getPineconeBaseUrl();
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
                agent: httpsAgent,
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
            const baseUrl = getPineconeBaseUrl();
            console.log('Fetching from Pinecone at URL:', baseUrl);
            
            const response = await fetch(`${baseUrl}/vectors/fetch`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Api-Key': process.env.PINECONE_API_KEY
                },
                body: JSON.stringify({
                    ids
                }),
                agent: httpsAgent,
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
    },

    testConnectivity: async () => {
        try {
            const baseUrl = getPineconeBaseUrl();
            console.log('Testing connectivity to Pinecone at:', baseUrl);
            
            const response = await fetch(`${baseUrl}/describe_index_stats`, {
                method: 'GET',
                headers: {
                    'Api-Key': process.env.PINECONE_API_KEY
                },
                agent: httpsAgent,
                timeout: 10000
            });
            
            if (response.ok) {
                const data = await response.json();
                console.log('Pinecone connectivity test successful!');
                console.log('Index stats:', JSON.stringify(data).substring(0, 200) + '...');
                return true;
            } else {
                console.error('Pinecone connectivity test failed with status:', response.status);
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
};

// Test connectivity on startup
let MOCK_MODE = false;
(async () => {
    try {
        const isConnected = await pineconeApi.testConnectivity();
        if (!isConnected) {
            console.warn('WARNING: Unable to connect to Pinecone. Switching to MOCK_MODE.');
            MOCK_MODE = true;
        } else {
            console.log('Successfully connected to Pinecone!');
        }
    } catch (error) {
        console.error('Error testing Pinecone connectivity:', error);
        console.warn('WARNING: Enabling MOCK_MODE due to connection error.');
        MOCK_MODE = true;
    }
})();

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

        // If in mock mode, just pretend we processed successfully
        if (MOCK_MODE) {
            console.log('MOCK MODE: Simulating successful processing');
            return {
                success: true,
                stats: {
                    totalItems: processedData.length,
                    processed: processedData.length,
                    skipped: 0,
                    errors: 0
                },
                sheets: workbook.worksheets.map(sheet => sheet.name),
                mockMode: true
            };
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
                    
                    // Try to fetch existing vector with error handling
                    let existingVector;
                    try {
                        const fetchResponse = await withRetry(() => pineconeApi.fetch([vectorId]));
                        existingVector = fetchResponse.vectors || {};
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
            
            // Upload the batch with retry logic
            if (batchOperations.length > 0) {
                try {
                    await withRetry(() => pineconeApi.upsert(batchOperations));
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
        // Special case for greetings
        const isGreeting = question.toLowerCase().match(/^(hi|hello|hey|greetings|howdy)[\s\.,!]*$/);
        if (isGreeting) {
            return {
                answer: "Hello! I'm your RFP Assistant. I can help you find information in your RFP documents. How can I assist you today?",
                sources: []
            };
        }

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

        // If in mock mode, just return a simulated response
        if (MOCK_MODE) {
            console.log('MOCK MODE: Simulating RFP query response');
            
            return {
                answer: `I'm currently operating in demo mode without database access. Your question was: "${question}". In normal operation, I would search our RFP database and provide relevant information based on the documents you've uploaded.`,
                sources: [],
                mockMode: true
            };
        }

        // Query Pinecone
        console.log('Querying Pinecone with embedding...');
        const queryResponse = await withRetry(() => 
            pineconeApi.query(queryEmbedding, 5, filterConditions)
        );
        
        console.log('Query response received:', !!queryResponse);

        // Handle potential empty responses
        if (!queryResponse.matches || queryResponse.matches.length === 0) {
            console.log('No matches found in Pinecone');
            return {
                answer: "I couldn't find any relevant information for your question in the RFP documents. Could you try rephrasing your question or ask about a different aspect of the RFP?",
                sources: []
            };
        }

        // Process the matches - FIX: Add error handling for JSON parsing
        const contexts = queryResponse.matches.map(match => {
            // Safe JSON parsing with error handling
            let parsedOriginalData = {};
            try {
                // Only attempt to parse if originalData exists and is a string
                if (match.metadata.originalData && typeof match.metadata.originalData === 'string') {
                    parsedOriginalData = JSON.parse(match.metadata.originalData);
                }
            } catch (error) {
                console.warn(`Error parsing originalData for match: ${error.message}`);
                // Continue with empty object if parsing fails
            }
            
            return {
                text: match.metadata.text || '',
                originalData: parsedOriginalData,
                category: match.metadata.category || 'unknown',
                sheetName: match.metadata.sheetName || 'unknown',
                similarity: match.score
            };
        });

        // Generate response using OpenAI
        const completion = await withRetry(() =>
            openai.chat.completions.create({
                model: "gpt-4",
                messages: [
                    {
                        role: "system",
                        content: `You are an RFP assistant specialized in analyzing historical RFP data.

For all responses:
- Be extremely concise and to the point
- Use direct language from the knowledge base whenever possible
- Only provide exactly what was asked, nothing more
- Do not offer explanations unless explicitly requested
- Format responses for quick reading and easy scanning

For general queries and greetings:
- Keep introductions minimal - identify as "RFP Assistant" only when first engaging
- Respond professionally but briefly
- Redirect to RFP topics if query is unrelated

For RFP-specific queries:
- Answer with precise information directly from the provided context
- Use the exact terminology and phrasing from the source documents
- If no clear answer exists in the knowledge base, state "I don't have that information" - do not attempt to extrapolate
- Never reference where information is coming from (no "according to..." or "as stated in...")
- Highlight only the most critical information requested
- If multiple interpretations of a question are possible, request clarification rather than guessing

Remember: Brevity is priority. Use minimal words to convey exact information.`
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
