function addToExcel(orderInfo) {
    let workbook;
    try {
      const workbookContent = fs.readFileSync('orders.xlsx');
      workbook = xlsx.read(workbookContent, { type: 'buffer' });
    } catch (error) {
      workbook = xlsx.utils.book_new();
    }
  
    const sheetName = `Order-${new Date().toISOString()}`;
    const worksheet = xlsx.utils.json_to_sheet([orderInfo]);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  
    const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    fs.writeFileSync('orders.xlsx', excelBuffer);
  }

  