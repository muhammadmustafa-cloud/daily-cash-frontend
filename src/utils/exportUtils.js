import { jsPDF } from 'jspdf'
import * as XLSX from 'xlsx'
// Import autotable function explicitly (ESM-friendly)
import autoTable from 'jspdf-autotable'

/**
 * Format date for display
 */
const formatDateForExport = (dateString) => {
  const date = new Date(dateString)
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
  const dayName = days[date.getDay()]
  const day = date.getDate().toString().padStart(2, '0')
  const month = (date.getMonth() + 1).toString().padStart(2, '0')
  const year = date.getFullYear()
  return { dayName, date: `${day}.${month}.${year}`, fullDate: date }
}

/**
 * Format currency for display
 */
const formatCurrencyForExport = (amount) => {
  return new Intl.NumberFormat('en-PK', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(amount || 0)
}


/**
 * Export Daily Cash Memo to PDF
 * Uses a single autoTable with side-by-side columns (ledger-style)
 */
export const exportToPDF = (memo, selectedDate, previousBalance) => {
  const creditEntries = memo?.creditEntries || []
  const debitEntries = memo?.debitEntries || []

  const { dayName, date: dateStr } = formatDateForExport(selectedDate)
  const doc = new jsPDF('p', 'mm', 'a4')

  // Ensure autotable is available (ESM import)
  if (!autoTable) {
    console.error('jspdf-autotable plugin not loaded')
    alert('PDF export plugin not available. Please refresh the page.')
    return
  }

  // ===== HEADER =====
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(14)
  doc.text('Daily Cash Memo', 105, 15, { align: 'center' })

  doc.setFontSize(10)
  doc.setFont('helvetica', 'normal')
  doc.text(`Date: ${dateStr}`, 14, 25)
  doc.text(`Day: ${dayName}`, 14, 32)

  // Section titles above table
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.text('CREDIT (Cash In)', 14, 40)
  doc.text('DEBIT (Cash Out)', 105, 40)

  // ===== PREPARE ROWS (Side-by-side format) =====
  const rows = []
  const maxRows = Math.max(creditEntries.length, debitEntries.length)

  // Data rows - align side by side
  for (let i = 0; i < maxRows; i++) {
    const c = creditEntries[i] || {}
    const d = debitEntries[i] || {}

    rows.push([
      c.name || '',
      c.description || '',
      c.amount ? formatCurrencyForExport(c.amount) : '',
      d.name || '',
      d.description || '',
      d.amount ? formatCurrencyForExport(d.amount) : ''
    ])
  }

  // Calculate totals
  const totalCredit = previousBalance + creditEntries.reduce((sum, e) => sum + (e.amount || 0), 0)
  const totalDebit = debitEntries.reduce((sum, e) => sum + (e.amount || 0), 0)

  // Total row
  rows.push([
    'Total', '', formatCurrencyForExport(totalCredit),
    'Total', '', formatCurrencyForExport(totalDebit)
  ])

  // Closing Balance row
  const closingBalance = totalCredit - totalDebit
  rows.push([
    '', { content: 'Closing Balance', styles: { fontStyle: 'bold', halign: 'center' } }, '',
    { content: formatCurrencyForExport(closingBalance), styles: { fontStyle: 'bold', halign: 'right' } }, '', ''
  ])

  // ===== TABLE (Single ledger-style table with 8 columns) =====
  // Add previous balance as a special row before the main table header
  autoTable(doc, {
    startY: 45,
    margin: { left: 14, right: 14 },
    body: [[
      { content: 'Previous Blnc', styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } },
      { content: '', styles: { fillColor: [240, 240, 240] } },
      { content: formatCurrencyForExport(previousBalance || 0), styles: { fontStyle: 'bold', halign: 'right', fillColor: [240, 240, 240] } },
      { content: 'Previous Blnc', styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } },
      { content: '', styles: { fillColor: [240, 240, 240] } },
      { content: '', styles: { fillColor: [240, 240, 240] } }
    ]],
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: 8.2,
      cellPadding: 1.8,
      valign: 'middle',
      lineWidth: 0.1,
      lineColor: [0, 0, 0]
    },
    columnStyles: {
      0: { cellWidth: 30, halign: 'left' },      // Credit Name
      1: { cellWidth: 40, halign: 'left' },      // Credit Description
      2: { cellWidth: 25, halign: 'right' },     // Credit Amount
      3: { cellWidth: 30, halign: 'left' },      // Debit Name
      4: { cellWidth: 40, halign: 'left' },      // Debit Description
      5: { cellWidth: 25, halign: 'right' }      // Debit Amount
    }
  })

  // Main table with headers
  autoTable(doc, {
    startY: doc.lastAutoTable.finalY + 2,
    margin: { left: 14, right: 14 },
    head: [[
      'Name', 'Description', 'Amount',
      'Name', 'Description', 'Amount'
    ]],
    body: rows,
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: 8.2,
      cellPadding: 1.8,
      valign: 'middle',
      lineWidth: 0.1,
      lineColor: [0, 0, 0]
    },
    headStyles: {
      fillColor: [240, 240, 240],
      textColor: [0, 0, 0],
      fontStyle: 'bold',
      fontSize: 9
    },
    bodyStyles: {
      textColor: [0, 0, 0]
    },
    columnStyles: {
      0: { cellWidth: 30, halign: 'left' },      // Credit Name
      1: { cellWidth: 40, halign: 'left' },      // Credit Description
      2: { cellWidth: 25, halign: 'right' },     // Credit Amount
      3: { cellWidth: 30, halign: 'left' },      // Debit Name
      4: { cellWidth: 40, halign: 'left' },      // Debit Description
      5: { cellWidth: 25, halign: 'right' }      // Debit Amount
    },
    bodyStyles: { overflow: 'linebreak' },
    // Style Total and Closing Balance rows
    didParseCell: function (data) {
      const rowIndex = data.row.index
      const bodyLength = rows.length
      
      // Check if it's the Total row (second to last) or Closing Balance row (last)
      if (rowIndex === bodyLength - 2 || rowIndex === bodyLength - 1) {
        data.cell.styles.fontStyle = 'bold'
      }
    }
  })

  // Save PDF
  const fileName = `Daily_Cash_Memo_${dateStr.replace(/\./g, '_')}.pdf`
  doc.save(fileName)
}

/**
 * Export Daily Cash Memo to Excel
 */
export const exportToExcel = (memo, selectedDate, previousBalance) => {
  // Allow export even if memo doesn't exist yet (will show empty entries)
  const creditEntries = memo?.creditEntries || []
  const debitEntries = memo?.debitEntries || []

  const { dayName, date: dateStr } = formatDateForExport(selectedDate)
  
  // Create a new workbook
  const wb = XLSX.utils.book_new()
  
  // CREDIT Sheet
  const creditData = [
    ['Daily Cash Memo'],
    ['Date', dateStr],
    ['Day', dayName],
    [],
    ['CREDIT (Cash In)'],
    ['Name', 'Description', 'Amount']
  ]
  
  // Previous Balance
  creditData.push(['Previous Blnc', '', previousBalance || 0])
  
  // Credit entries
  if (creditEntries.length > 0) {
    creditEntries.forEach(entry => {
      creditData.push([
        entry.name,
        entry.description || '',
        entry.amount || 0
      ])
    })
  }
  
  // Total Credit
  const totalCredit = previousBalance + creditEntries.reduce((sum, e) => sum + (e.amount || 0), 0)
  creditData.push(['Total', '', totalCredit])
  
  const creditWs = XLSX.utils.aoa_to_sheet(creditData)
  
  // Set column widths
  creditWs['!cols'] = [
    { wch: 25 }, // Name
    { wch: 30 }, // Description
    { wch: 15 }  // Amount
  ]
  
  XLSX.utils.book_append_sheet(wb, creditWs, 'Credit')
  
  // DEBIT Sheet
  const debitData = [
    ['Daily Cash Memo'],
    ['Date', dateStr],
    ['Day', dayName],
    [],
    ['DEBIT (Cash Out)'],
    ['Name', 'Description', 'Amount']
  ]
  
  // Previous Balance (empty for debit)
  debitData.push(['Previous Blnc', '', ''])
  
  // Debit entries
  if (debitEntries.length > 0) {
    debitEntries.forEach(entry => {
      debitData.push([
        entry.name,
        entry.description || '',
        entry.amount || 0
      ])
    })
  }
  
  // Total Debit
  const totalDebit = debitEntries.reduce((sum, e) => sum + (e.amount || 0), 0)
  debitData.push(['Total', '', totalDebit])
  
  const debitWs = XLSX.utils.aoa_to_sheet(debitData)
  
  // Set column widths
  debitWs['!cols'] = [
    { wch: 25 }, // Name
    { wch: 30 }, // Description
    { wch: 15 }  // Amount
  ]
  
  XLSX.utils.book_append_sheet(wb, debitWs, 'Debit')
  
  // Combined Sheet - Side by Side Layout
  const combinedData = [
    ['Daily Cash Memo', '', '', '', '', 'Daily Cash Memo'],
    ['Date', dateStr, '', '', '', 'Date', dateStr],
    ['Day', dayName, '', '', '', 'Day', dayName],
    [],
    ['CREDIT (Cash In)', '', '', '', '', 'DEBIT (Cash Out)', '', ''],
    ['Name', 'Description', 'Amount', '', '', 'Name', 'Description', 'Amount']
  ]
  
  // Prepare credit and debit rows for side-by-side layout
  const creditRows = []
  creditRows.push(['Previous Blnc', '', previousBalance || 0])
  if (creditEntries.length > 0) {
    creditEntries.forEach(entry => {
      creditRows.push([
        entry.name,
        entry.description || '',
        entry.amount || 0
      ])
    })
  }
  creditRows.push(['Total', '', totalCredit])
  
  const debitRows = []
  debitRows.push(['Previous Blnc', '', ''])
  if (debitEntries.length > 0) {
    debitEntries.forEach(entry => {
      debitRows.push([
        entry.name,
        entry.description || '',
        entry.amount || 0
      ])
    })
  }
  debitRows.push(['Total', '', totalDebit])
  
  // Combine credit and debit rows side by side
  const maxRows = Math.max(creditRows.length, debitRows.length)
  for (let i = 0; i < maxRows; i++) {
    const creditRow = creditRows[i] || ['', '', '']
    const debitRow = debitRows[i] || ['', '', '']
    combinedData.push([
      creditRow[0] || '', // Credit Name
      creditRow[1] || '', // Credit Description
      creditRow[2] !== undefined ? creditRow[2] : '', // Credit Amount
      '', '', // Empty columns for spacing
      debitRow[0] || '', // Debit Name
      debitRow[1] || '', // Debit Description
      debitRow[2] !== undefined ? debitRow[2] : ''  // Debit Amount
    ])
  }
  
  combinedData.push([])
  combinedData.push(['Closing Balance', '', totalCredit - totalDebit, '', '', '', '', ''])
  
  const combinedWs = XLSX.utils.aoa_to_sheet(combinedData)
  combinedWs['!cols'] = [
    { wch: 25 }, // Credit Name
    { wch: 30 }, // Credit Description
    { wch: 15 }, // Credit Amount
    { wch: 5 },  // Spacing
    { wch: 5 },  // Spacing
    { wch: 25 }, // Debit Name
    { wch: 30 }, // Debit Description
    { wch: 15 }  // Debit Amount
  ]
  
  XLSX.utils.book_append_sheet(wb, combinedWs, 'Combined')
  
  // Save Excel file
  const fileName = `Daily_Cash_Memo_${dateStr.replace(/\./g, '_')}.xlsx`
  XLSX.writeFile(wb, fileName)
}

