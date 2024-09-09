const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware to parse JSON bodies
app.use(express.json());

// Basic route
app.get('/', (req, res) => {
    res.send('Hello, World!');
});
app.get('/about', (req, res) => {
    res.send('About this application');
});
app.get('/cats', (req, res) => {
    res.send('About this cats');
});
app.get('/generate-word', async (req, res) => {
    try {
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({
                        text: "Hello World",
                        heading: "Title",
                    }),
                    new Paragraph({
                        children: [
                            new TextRun("This is a simple Word document generated using Node.js."),
                        ],
                    }),
                ],
            }],
        });

        const buffer = await Packer.toBuffer(doc);

        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=generated.docx',
        });

        res.send(buffer);
    } catch (error) {
        console.error(error);
        res.status(500).send('Error generating document');
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});