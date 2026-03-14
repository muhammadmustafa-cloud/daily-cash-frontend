import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Search, Edit2, Trash2, Package, Download } from 'lucide-react'
import Card, { CardContent, CardDescription, CardHeader, CardTitle } from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import { FormInput } from '../../components/ui/form-input'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '../../components/ui/dialog'
import { DataTable } from '../../components/ui/data-table'
import api from '../../config/api'
import { formatDate } from '../../utils/formatters'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'

const LabourRates = () => {
  const [labourRates, setLabourRates] = useState([])
  const [labourExpenses, setLabourExpenses] = useState([])
  const [loading, setLoading] = useState(false)
  const [loadingExpenses, setLoadingExpenses] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingLabourRate, setEditingLabourRate] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [saving, setSaving] = useState(false)
  const [filters, setFilters] = useState({ startDate: '', endDate: '' })

  const { register, handleSubmit, formState: { errors }, reset, setValue } = useForm()

  // Fetch expenses once on mount (does not depend on search)
  useEffect(() => {
    fetchLabourExpenses()
  }, [])

  // Initial load
  useEffect(() => {
    fetchLabourRates({ search: searchTerm, startDate: filters.startDate, endDate: filters.endDate })
  }, [])

  const fetchLabourRates = async ({ search = '', startDate = '', endDate = '' } = {}) => {
    try {
      setLoading(true)
      const today = new Date().toISOString().slice(0, 10)
      const effectiveStart = startDate || endDate || today
      const effectiveEnd = endDate || startDate || today
      const response = await api.get('/labour-rates', {
        params: {
          page: 1,
          limit: 1000,
          search,
          startDate: effectiveStart,
          endDate: effectiveEnd,
          isActive: true,
        },
      })
      setLabourRates(response.data.data || [])
    } catch (error) {
      toast.error('Failed to fetch labour records')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const fetchLabourExpenses = async () => {
    try {
      setLoadingExpenses(true)
      const response = await api.get('/labour-expenses', {
        params: { page: 1, limit: 1000, isActive: true }
      })
      setLabourExpenses(response.data.data || [])
    } catch (error) {
      toast.error('Failed to fetch labour expenses')
      console.error(error)
    } finally {
      setLoadingExpenses(false)
    }
  }

  const handleAddLabourRate = () => {
    setEditingLabourRate(null)
    reset({
      labourExpense: '',
      bags: ''
    })
    setIsModalOpen(true)
  }

  const handleEditLabourRate = (labourRate) => {
    setEditingLabourRate(labourRate)
    setValue('labourExpense', labourRate.labourExpense?._id || '')
    setValue('bags', labourRate.bags)
    setIsModalOpen(true)
  }

  const handleDeleteLabourRate = async (row) => {
    if (!window.confirm('Are you sure you want to delete this labour record?')) {
      return
    }

    const id = row?._id || row?.id
    if (!id || typeof id !== 'string') {
      toast.error('Invalid labour record')
      return
    }

    try {
      await api.delete(`/labour-rates/${id}`)
      toast.success('Labour record deleted successfully')
      fetchLabourRates({ search: searchTerm, startDate: filters.startDate, endDate: filters.endDate })
    } catch (error) {
      toast.error('Failed to delete labour record')
      console.error(error)
    }
  }

  const onSubmit = async (data) => {
    try {
      setSaving(true)
      // Coerce and validate inputs
      const selectedExpense = labourExpenses.find(le => le._id === data.labourExpense)
      const bags = Number(data.bags)

      if (!selectedExpense) {
        toast.error('Please select a valid labour')
        return
      }
      if (Number.isNaN(bags) || bags < 0) {
        toast.error('Bags must be 0 or more')
        return
      }

      const payload = {
        labourExpense: data.labourExpense,
        bags,
        rate: Number(selectedExpense.rate) || 0,
      }

      if (editingLabourRate) {
        await api.put(`/labour-rates/${editingLabourRate._id}`, payload)
        toast.success('Labour record updated successfully')
      } else {
        await api.post('/labour-rates', payload)
        toast.success('Labour record created successfully')
      }
      setIsModalOpen(false)
      fetchLabourRates({ search: searchTerm, startDate: filters.startDate, endDate: filters.endDate })
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save labour record')
      console.error(error)
    } finally {
      setSaving(false)
    }
  }

  const handleDownloadPDF = () => {
    try {
      const doc = new jsPDF()
      
      // Add custom font for better Unicode support
      doc.setFont('helvetica')
      
      // Header with better styling
      doc.setFontSize(24)
      doc.setTextColor(0, 0, 0)
      doc.text('Muslim Daal Mill', 105, 25, { align: 'center' })
      
      doc.setFontSize(18)
      doc.setTextColor(59, 130, 246)
      doc.text('Labour Records Report', 105, 35, { align: 'center' })
      
      // Add a line under the title
      doc.setDrawColor(59, 130, 246)
      doc.setLineWidth(0.5)
      doc.line(20, 40, 190, 40)
      
      // Report info with better positioning
      doc.setFontSize(10)
      doc.setTextColor(100, 100, 100)
      const reportDate = new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      })
      doc.text(`Generated on: ${reportDate}`, 105, 50, { align: 'center' })
      
      // Calculate totals
      const totalBags = labourRates.reduce((sum, lr) => sum + (Number(lr.bags) || 0), 0)
      const totalAmount = labourRates.reduce((sum, lr) => sum + ((Number(lr.bags) || 0) * (Number(lr.rate) || 0)), 0)
      
      // Table with better styling and totals
      const tableData = labourRates.map(lr => [
        lr.labourExpense?.name || 'Unknown',
        (Number(lr.bags) || 0).toString(),
        `PKR ${Number(lr.rate || 0).toFixed(2)}`,
        `PKR ${(((Number(lr.bags) || 0) * (Number(lr.rate) || 0))).toFixed(2)}`,
        formatDate(lr.createdAt)
      ])
      
      // Add total row
      tableData.push([
        'TOTAL',
        totalBags.toString(),
        '',
        `PKR ${totalAmount.toFixed(2)}`,
        ''
      ])
      
      autoTable(doc, {
        head: [['Labour Name', 'Bags', 'Rate', 'Total Amount', 'Created Date']],
        body: tableData,
        startY: 65,
        theme: 'plain',
        styles: {
          fontSize: 10,
          cellPadding: 4,
          lineColor: [200, 200, 200],
          textColor: [40, 40, 40]
        },
        headStyles: {
          fillColor: [245, 245, 245],
          textColor: [0, 0, 0],
          fontStyle: 'bold',
          fontSize: 10,
        },
        columnStyles: {
          0: { cellWidth: 55, halign: 'left' }, // Labour Name
          1: { cellWidth: 25, halign: 'center' }, // Bags
          2: { cellWidth: 35, halign: 'right' }, // Rate
          3: { cellWidth: 40, halign: 'right' }, // Total Amount
          4: { cellWidth: 35, halign: 'left' } // Created Date
        },
        // Style the last row (total row) differently
        didParseCell: (data) => {
          if (data.row.section === 'body' && data.row.index === tableData.length - 1) {
            data.cell.styles.fontStyle = 'bold'
            data.cell.styles.fillColor = [250, 250, 250]
            data.cell.styles.textColor = [0, 0, 0]
            if (data.column.index === 0) {
              data.cell.styles.halign = 'left'
            } else if (data.column.index === 1 || data.column.index === 3) {
              data.cell.styles.halign = 'right'
            }
          }
        },
        margin: { top: 65, left: 20, right: 20 }
      })
      
      // Footer with better styling
      const pageCount = doc.internal.getNumberOfPages()
      doc.setFontSize(9)
      doc.setTextColor(150, 150, 150)
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i)
        
        // Add footer line
        doc.setDrawColor(200, 200, 200)
        doc.setLineWidth(0.3)
        doc.line(20, doc.internal.pageSize.height - 15, 190, doc.internal.pageSize.height - 15)
        
        doc.text(`Page ${i} of ${pageCount}`, 105, doc.internal.pageSize.height - 10, { align: 'center' })
        doc.text('Muslim Daal Mill Management System', 105, doc.internal.pageSize.height - 5, { align: 'center' })
      }
      
      // Save the PDF with better filename
      const fileName = `Labour-Records-${new Date().toISOString().split('T')[0]}.pdf`
      doc.save(fileName)
      
      toast.success('PDF downloaded successfully')
    } catch (error) {
      console.error('Error generating PDF:', error)
      toast.error('Failed to generate PDF')
    }
  }

  const columns = [
    {
      key: 'labourExpense.name',
      label: 'Labour Name',
      render: (value, row) => (
        <div className="font-medium text-gray-900">{row.labourExpense?.name || 'Unknown'}</div>
      ),
    },
    {
      key: 'bags',
      label: 'Bags',
      render: (value, row) => (
        <div className="flex items-center">
          <Package className="h-4 w-4 text-gray-500 mr-1" />
          <span className="font-medium">{Number(value) || 0}</span>
        </div>
      ),
    },
    {
      key: 'rate',
      label: 'Rate',
      render: (value, row) => (
        <div className="flex items-center">
          <span className="text-sm font-bold text-green-600 mr-1">PKR</span>
          <span className="font-medium">{Number(value || 0).toFixed(2)}</span>
        </div>
      ),
    },
    {
      key: 'totalAmount',
      label: 'Total Amount',
      render: (value, row) => {
        const total = (Number(row.bags) || 0) * (Number(row.rate) || 0)
        return (
          <div className="flex items-center">
            <span className="text-sm font-bold text-green-600 mr-1">PKR</span>
            <span className="font-semibold text-green-600">{total.toFixed(2)}</span>
          </div>
        )
      },
    },
    {
      key: 'isActive',
      label: 'Status',
      render: (value, row) => (
        <span
          className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${
            value
              ? 'bg-green-100 text-green-800'
              : 'bg-red-100 text-red-800'
          }`}
        >
          {value ? 'Active' : 'Inactive'}
        </span>
      ),
    },
    {
      key: 'createdAt',
      label: 'Created',
      render: (value, row) => formatDate(value),
    },
  ]

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Labour</h1>
          <p className="text-gray-600">Manage labour work records and bag counts</p>
        </div>
        <div className="flex items-center gap-3">
          <Button 
            onClick={handleDownloadPDF}
            variant="outline"
            disabled={labourRates.length === 0}
          >
            <Download className="h-4 w-4 mr-2" />
            Download PDF
          </Button>
          <Button onClick={handleAddLabourRate}>
            <Plus className="h-4 w-4 mr-2" />
            Add Labour Record
          </Button>
        </div>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>Labour Records</CardTitle>
          <CardDescription>
            View and manage all labour work records in your system
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6 items-end">
            <div className="w-full">
              <label className="block text-sm font-medium text-gray-700 mb-1">Start Date</label>
              <input
                type="date"
                value={filters.startDate}
                onChange={(e) => setFilters(prev => ({ ...prev, startDate: e.target.value }))}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent"
              />
            </div>
            <div className="w-full">
              <label className="block text-sm font-medium text-gray-700 mb-1">End Date</label>
              <input
                type="date"
                value={filters.endDate}
                onChange={(e) => setFilters(prev => ({ ...prev, endDate: e.target.value }))}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent"
              />
            </div>
            <div className="w-full relative">
              <label className="block text-sm font-medium text-gray-700 mb-1">Search</label>
              <Search className="absolute left-3 top-[42px] text-gray-400 h-4 w-4" />
              <input
                type="text"
                placeholder="Search labour records..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent"
              />
            </div>
            <div className="w-full">
              <label className="block text-sm font-medium text-gray-700 mb-1">&nbsp;</label>
              <Button
                onClick={() => fetchLabourRates({ search: searchTerm, startDate: filters.startDate, endDate: filters.endDate })}
                className="w-full"
                disabled={loading}
              >
                Apply Filters
              </Button>
            </div>
          </div>

          <DataTable
            columns={columns}
            data={labourRates}
            loading={loading}
            onEdit={handleEditLabourRate}
            onDelete={handleDeleteLabourRate}
          />
        </CardContent>
      </Card>

      <Dialog open={isModalOpen} onOpenChange={setIsModalOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>
              {editingLabourRate ? 'Edit Labour Record' : 'Add Labour Record'}
            </DialogTitle>
            <DialogDescription>
              {editingLabourRate
                ? 'Update labour record information.'
                : 'Add a new labour record to your system.'}
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(onSubmit)}>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Labour Name
                </label>
                <select
                  disabled={loadingExpenses}
                  {...register('labourExpense', { required: 'Labour name is required' })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent"
                >
                  <option value="" disabled={loadingExpenses}>
                    {loadingExpenses ? 'Loading labour list...' : 'Select Labour'}
                  </option>
                  {labourExpenses.map((expense) => (
                    <option key={expense._id} value={expense._id}>
                      {expense.name} (PKR {Number(expense.rate || 0).toFixed(2)})
                    </option>
                  ))}
                </select>
                {errors.labourExpense && (
                  <p className="text-red-500 text-sm mt-1">{errors.labourExpense.message}</p>
                )}
              </div>

              <FormInput
                label="Bags"
                name="bags"
                type="number"
                step="1"
                min="0"
                register={register}
                errors={errors}
                placeholder="Enter number of bags"
                required
              />
            </div>
            <DialogFooter>
              <Button
                type="button"
                variant="outline"
                onClick={() => setIsModalOpen(false)}
              >
                Cancel
              </Button>
              <Button type="submit" disabled={saving}>
                {saving ? 'Saving...' : (editingLabourRate ? 'Update' : 'Create')}
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}

export default LabourRates
