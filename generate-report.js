const ExcelJS = require('exceljs');
const fs = require('fs');

async function generateReport() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Test Results');
  worksheet.columns = [
    { header: 'Test Case', key: 'testCase', width: 30 },
    { header: 'Status', key: 'status', width: 10 },
    { header: 'Duration', key: 'duration', width: 10 },
  ];

  // Check if the results.json file exists
  if (!fs.existsSync('cypress/results.json')) {
    console.error('cypress/results.json file not found');
    process.exit(1);
  }

  // Read Cypress test results (assuming JSON format)
  const results = JSON.parse(fs.readFileSync('cypress/results.json', 'utf8'));
  console.log('Test results:', results); // Add this line for debugging
  results.forEach(result => {
    worksheet.addRow({
      testCase: result.testCase,
      status: result.status,
      duration: result.duration,
    });
  });

  await workbook.xlsx.writeFile('cypress/results.xlsx');
  console.log('Excel report generated successfully.');
}

generateReport().catch(console.error);