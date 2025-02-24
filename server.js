// Load environment variables
require('dotenv').config();

// Import dependencies
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { processExcelRFP, queryRFPData } = require('./excelProcessor');

// Verify environment variables
console.log('Environment variables check:');
console.log('OPENAI_API_KEY exists:', !!process.env.OPENAI_API_KEY);
console.log('PINECONE_API_KEY exists:', !!process.env.PINECONE_API_KEY);
console.log('PINECONE_ENVIRONMENT exists:', !!process.env.PINECONE_ENVIRONMENT);
console.log('PINECONE_INDEX_NAME exists:', !!process.env.PINECONE_INDEX_NAME);

// Initialize Express app
const app = express();

// Configure CORS
// In production, replace with your actual Render URL
const FRONTEND_URL = process.env.NODE_ENV === 'production' 
  ? 'https://your-render-frontend-url.onrender.com'
  : 'http://localhost:3000';

app.use(cors({
    origin: [FRONTEND_URL, 'http://localhost:3000'],
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true
}));

// Middleware
app.use(express.json());

// Configure multer for file uploads
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
    destination: uploadDir,
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const upload = multer({
    storage,
    fileFilter: (req, file, cb) => {
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        if (allowedTypes.includes(file.mimetype) || 
            file.originalname.match(/\.(xlsx|xls)$/)) {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files are allowed!'), false);
        }
    },
    limits: {
        fileSize: 5 * 1024 * 1024 // 5MB limit
    }
});

// API Routes
app.post('/api/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const metadata = {
            rfpId: req.body.rfpId || `RFP-${Date.now()}`,
            title: req.body.title || req.file.originalname,
            uploadDate: new Date().toISOString(),
            category: req.body.category || 'uncategorized'
        };

        const fileBuffer = fs.readFileSync(req.file.path);
        const result = await processExcelRFP(fileBuffer, metadata);

        // Clean up uploaded file
        fs.unlinkSync(req.file.path);

        res.json({
            success: true,
            message: 'Excel RFP processed successfully',
            details: result
        });
    } catch (error) {
        console.error('Error processing upload:', error);
        res.status(500).json({ 
            error: error.message || 'Error processing file',
            details: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

app.post('/api/query', async (req, res) => {
    try {
        console.log('Received query:', req.body.question);
        const { question, filters } = req.body;
        const response = await queryRFPData(question, filters);
        res.json(response);
    } catch (error) {
        console.error('Error processing query:', error);
        res.status(500).json({ 
            error: error.message || 'Error processing query',
            details: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        environment: process.env.NODE_ENV,
        timestamp: new Date().toISOString()
    });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ 
        error: 'Something broke!',
        details: process.env.NODE_ENV === 'development' ? err.stack : undefined
    });
});

// Start server
const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running in ${process.env.NODE_ENV || 'development'} mode on port ${PORT}`);
    console.log(`Health check available at: http://localhost:${PORT}/health`);
});

// Handle uncaught exceptions
process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
    // Implement your error notification service here if needed
});

// Handle unhandled promise rejections
process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
    // Implement your error notification service here if needed
});