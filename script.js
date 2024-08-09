document.getElementById('billForm').addEventListener('submit', function(event) {
    event.preventDefault();
    
    const submissionDate = document.getElementById('submissionDate').value;
    const amount = parseFloat(document.getElementById('amount').value);
    const purpose = document.getElementById('purpose').value;
    const description = document.getElementById('description').value;
    const projectNumber = document.getElementById('projectNumber').value;
    const payeeName = document.getElementById('payeeName').value;
    const uploadBill = document.getElementById('uploadBill').files[0];

    if (!submissionDate || !amount || !purpose || !description || !projectNumber || !payeeName || !uploadBill) {
        document.getElementById('message').innerText = 'Please fill all fields and upload a bill.';
        return;
    }

    // Create and display message
    document.getElementById('message').innerText = 'Bill submitted successfully!';

    // Save data to Excel
    saveToExcel(submissionDate, amount, purpose, description, projectNumber, payeeName);

    // Reset the form
    document.getElementById('billForm').reset();
});

// Function to save data to Excel
function saveToExcel(submissionDate, amount, purpose, description, projectNumber, payeeName) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet([{
        SubmissionDate: submissionDate,
        Amount: amount,
        Purpose: purpose,
        Description: description,
        ProjectNumber: projectNumber,
        PayeeName: payeeName
    }]);

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Bills');

    // Save to file
    XLSX.writeFile(workbook, 'bills.xlsx');
}

// Include the XLSX library for Excel file creation
document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"><\/script>');
