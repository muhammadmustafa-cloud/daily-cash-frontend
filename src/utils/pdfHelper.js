// Helper file to ensure jspdf-autotable is loaded
import { jsPDF } from 'jspdf'

// Import jspdf-autotable - this should extend jsPDF prototype
// Using dynamic import to ensure it loads
let pluginLoaded = false

export const ensureAutoTableLoaded = async () => {
  if (pluginLoaded) return true
  
  try {
    // Try side-effect import first
    await import('jspdf-autotable')
    pluginLoaded = true
    return true
  } catch (error) {
    console.error('Failed to load jspdf-autotable:', error)
    return false
  }
}

export const createPDF = () => {
  return new jsPDF()
}

