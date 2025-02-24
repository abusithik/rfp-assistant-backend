// server/excelProcessor.js
// Check required environment variables
const requiredEnvVars = ['PINECONE_API_KEY', 'PINECONE_ENVIRONMENT', 'PINECONE_INDEX_NAME'];
const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingVars.length > 0) {
    throw new Error(`Missing required environment variables: ${missingVars.join(', ')}`);
}

console.log('Debug - Environment Variables:');
console.log('PINECONE_API_KEY exists:', !!process.env.PINECONE_API_KEY);
console.log('PINECONE_ENVIRONMENT exists:', !!process.env.PINECONE_ENVIRONMENT);
console.log('PINECONE_ENVIRONMENT value:', process.env.PINECONE_ENVIRONMENT);
console.log('PINECONE_INDEX_NAME exists:', !!process.env.PINECONE_INDEX_NAME);

console.log('Pinecone Configuration:', {
    environment: process.env.PINECONE_ENVIRONMENT,
    indexName: process.env.PINECONE_INDEX_NAME
});

const ExcelJS = require('exceljs');
const _ = require('lodash');
const { Pinecone } = require('@pinecone-database/pinecone');
const { OpenAI } = require('openai');
const fetch = require('node-fetch');
const https = require('https');
const crypto = require('crypto');

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Update customFetch configuration
const customFetch = (url, options = {}) => {
    return fetch(url, {
        ...options,
        agent: new https.Agent({
            rejectUnauthorized: false,
            keepAlive: true,
            timeout: 60000
        }),
        timeout: 60000,
        headers: {
            ...options.headers,
            'Api-Key': process.env.PINECONE_API_KEY
        }
    });
};

const pinecone = new Pinecone({
    apiKey: process.env.PINECONE_API_KEY,
    environment: process.env.PINECONE_ENVIRONMENT
});

const index = pinecone.index(process.env.PINECONE_INDEX_NAME);

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
          
          // Try to fetch existing vector with error handling
          let existingVector;
          try {
            const fetchResponse = await index.fetch([vectorId]);
            existingVector = fetchResponse.vectors || {};
          } catch (fetchError) {
            console.log(`Fetch check failed for ${vectorId}, proceeding with upsert`);
            existingVector = {};
          }
          
          if (!existingVector[vectorId]) {
            const embedding = await getEmbedding(item.text);
            
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
          continue; // Continue with next item even if this one fails
        }
      }
      
      // Upload the batch
      if (batchOperations.length > 0) {
        try {
          await index.upsert(batchOperations);
          console.log(`Successfully uploaded batch of ${batchOperations.length} items`);
        } catch (batchError) {
          console.error(`Error uploading batch: ${batchError.message}`);
          errorRows += batchOperations.length;
          // Could implement retry logic here if needed
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
    const queryEmbedding = await getEmbedding(question);

    // Prepare filter conditions - add a default filter if none provided
    let filterConditions = {};
    if (Object.keys(filters).length > 0) {
      // Use provided filters
      if (filters.category) filterConditions.category = filters.category;
      if (filters.sheetName) filterConditions.sheetName = filters.sheetName;
    } else {
      // Add a default filter - filter for any non-null category
      filterConditions = {
        category: { $exists: true }
      };
    }

    // Query Pinecone with query parameters
    const queryResponse = await index.query({
      vector: queryEmbedding,
      topK: 5,
      includeMetadata: true,
      ...(Object.keys(filterConditions).length > 0 && { filter: filterConditions })
    });

    // Process and format results
    const contexts = queryResponse.matches.map(match => ({
      text: match.metadata.text,
      originalData: JSON.parse(match.metadata.originalData),
      category: match.metadata.category,
      sheetName: match.metadata.sheetName,
      similarity: match.score
    }));

    // Generate response using OpenAI
    const completion = await openai.chat.completions.create({
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
    });

    return {
      answer: completion.choices[0].message.content,
      sources: contexts
    };
  } catch (error) {
    console.error('Error querying RFP data:', error);
    throw error;
  }
}

module.exports = {
  processExcelRFP,
  queryRFPData
};
