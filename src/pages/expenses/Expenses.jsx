import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Edit, Trash2, Search } from 'lucide-react'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Select from '../../components/ui/Select'
import Modal from '../../components/ui/Modal'
import Table from '../../components/ui/Table'
import api from '../../config/api'
import { formatCurrency, formatDate } from '../../utils/formatters'

const Expenses = () => {
  const [expenses, setExpenses] = useState([])
  const [mazdoors, setMazdoors] = useState([])
  const [suppliers, setSuppliers] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingExpense, setEditingExpense] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [pagination, setPagination] = useState({ page: 1, limit: 10, total: 0 })

  const { register, handleSubmit, formState: { errors }, reset, setValue, watch } = useForm()
  const category = watch('category')

  useEffect(() => {
    fetchExpenses()
    fetchMazdoors()
    fetchSuppliers()
  }, [pagination.page, searchTerm])

  const fetchMazdoors = async () => {
    try {
      const response = await api.get('/mazdoors?limit=1000')
      setMazdoors(response.data.data || [])
    } catch (error) {
      console.error('Error fetching mazdoors:', error)
    }
  }

  const fetchSuppliers = async () => {
    try {
      const response = await api.get('/suppliers?limit=1000')
      setSuppliers(response.data.data || [])
    } catch (error) {
      console.error('Error fetching suppliers:', error)
    }
  }

  const fetchExpenses = async () => {
    try {
      setLoading(true)
      const response = await api.get('/expenses', {
        params: {
          page: pagination.page,
          limit: pagination.limit,
          search: searchTerm,
        },
      })
      setExpenses(response.data.data || [])
      setPagination(prev => ({
        ...prev,
        total: response.data.pagination?.total || 0,
      }))
    } catch (error) {
      toast.error('Failed to fetch expenses')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    setEditingExpense(null)
    reset()
    setValue('date', new Date().toISOString().split('T')[0])
    setValue('paymentMethod', 'cash')
    setIsModalOpen(true)
  }

  const handleEdit = (expense) => {
    setEditingExpense(expense)
    setValue('category', expense.category)
    setValue('subCategory', expense.subCategory || '')
    setValue('description', expense.description)
    setValue('amount', expense.amount)
    setValue('date', expense.date ? expense.date.split('T')[0] : new Date().toISOString().split('T')[0])
    setValue('paymentMethod', expense.paymentMethod || 'cash')
    setValue('mazdoor', expense.mazdoor?._id || '')
    setValue('supplier', expense.supplier?._id || '')
    setValue('billNumber', expense.billNumber || '')
    setValue('notes', expense.notes || '')
    setIsModalOpen(true)
  }

  const handleDelete = async (expense) => {
    if (!window.confirm(`Are you sure you want to delete this expense?`)) {
      return
    }

    try {
      await api.delete(`/expenses/${expense._id}`)
      toast.success('Expense deleted successfully')
      fetchExpenses()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete expense')
    }
  }

  const onSubmit = async (data) => {
    try {
      if (editingExpense) {
        await api.put(`/expenses/${editingExpense._id}`, data)
        toast.success('Expense updated successfully')
      } else {
        await api.post('/expenses', data)
        toast.success('Expense created successfully')
      }
      setIsModalOpen(false)
      reset()
      fetchExpenses()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save expense')
    }
  }

  const columns = [
    {
      key: 'date',
      label: 'Date',
      render: (value) => formatDate(value),
    },
    {
      key: 'category',
      label: 'Category',
      render: (value) => value?.replace('_', ' ').toUpperCase() || '-',
    },
    { key: 'description', label: 'Description' },
    {
      key: 'amount',
      label: 'Amount',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'mazdoor',
      label: 'Mazdoor',
      render: (value) => value?.name || '-',
    },
    {
      key: 'supplier',
      label: 'Supplier',
      render: (value) => value?.name || '-',
    },
    { key: 'paymentMethod', label: 'Payment Method' },
  ]

  const actions = [
    { key: 'edit', label: 'Edit', variant: 'primary' },
    { key: 'delete', label: 'Delete', variant: 'danger' },
  ]

  const handleAction = (action, expense) => {
    if (action === 'edit') {
      handleEdit(expense)
    } else if (action === 'delete') {
      handleDelete(expense)
    }
  }

  const categoryOptions = [
    { value: 'mazdoor', label: 'Mazdoor' },
    { value: 'electricity', label: 'Electricity' },
    { value: 'rent', label: 'Rent' },
    { value: 'transport', label: 'Transport' },
    { value: 'maintenance', label: 'Maintenance' },
    { value: 'raw_material', label: 'Raw Material' },
    { value: 'packaging', label: 'Packaging' },
    { value: 'other', label: 'Other' },
  ]

  const paymentMethodOptions = [
    { value: 'cash', label: 'Cash' },
    { value: 'bank', label: 'Bank' },
    { value: 'cheque', label: 'Cheque' },
    { value: 'online', label: 'Online' },
  ]

  const mazdoorOptions = [
    { value: '', label: 'None' },
    ...mazdoors.map(m => ({ value: m._id, label: m.name })),
  ]

  const supplierOptions = [
    { value: '', label: 'None' },
    ...suppliers.map(s => ({ value: s._id, label: s.name })),
  ]

  return (
    <div>
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Expenses</h1>
          <p className="text-gray-600 mt-1">Manage all expenses</p>
        </div>
        <Button onClick={handleCreate} variant="primary">
          <Plus className="h-5 w-5 mr-2" />
          Add Expense
        </Button>
      </div>

      <Card>
        <div className="mb-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search expenses..."
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value)
                setPagination(prev => ({ ...prev, page: 1 }))
              }}
              className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
            />
          </div>
        </div>

        {loading ? (
          <div className="flex justify-center py-8">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary-600"></div>
          </div>
        ) : (
          <Table
            columns={columns}
            data={expenses}
            actions={actions}
            onAction={handleAction}
          />
        )}

        {pagination.total > pagination.limit && (
          <div className="mt-4 flex justify-between items-center">
            <p className="text-sm text-gray-600">
              Showing {((pagination.page - 1) * pagination.limit) + 1} to{' '}
              {Math.min(pagination.page * pagination.limit, pagination.total)} of{' '}
              {pagination.total} expenses
            </p>
            <div className="flex gap-2">
              <Button
                variant="outline"
                size="sm"
                disabled={pagination.page === 1}
                onClick={() => setPagination(prev => ({ ...prev, page: prev.page - 1 }))}
              >
                Previous
              </Button>
              <Button
                variant="outline"
                size="sm"
                disabled={pagination.page * pagination.limit >= pagination.total}
                onClick={() => setPagination(prev => ({ ...prev, page: prev.page + 1 }))}
              >
                Next
              </Button>
            </div>
          </div>
        )}
      </Card>

      <Modal
        isOpen={isModalOpen}
        onClose={() => {
          setIsModalOpen(false)
          reset()
        }}
        title={editingExpense ? 'Edit Expense' : 'Add New Expense'}
        size="md"
      >
        <form onSubmit={handleSubmit(onSubmit)}>
          <Select
            label="Category"
            name="category"
            register={register}
            required
            options={categoryOptions}
            error={errors.category?.message}
          />
          <Input
            label="Sub Category"
            name="subCategory"
            register={register}
            error={errors.subCategory?.message}
            placeholder="Optional"
          />
          <Input
            label="Description"
            name="description"
            register={register}
            required
            error={errors.description?.message}
            placeholder="Expense description"
          />
          <div className="grid grid-cols-2 gap-4">
            <Input
              label="Amount"
              name="amount"
              type="number"
              step="0.01"
              register={register}
              required
              error={errors.amount?.message}
              placeholder="0.00"
            />
            <Input
              label="Date"
              name="date"
              type="date"
              register={register}
              required
              error={errors.date?.message}
            />
          </div>
          <Select
            label="Payment Method"
            name="paymentMethod"
            register={register}
            options={paymentMethodOptions}
            error={errors.paymentMethod?.message}
          />
          {category === 'mazdoor' && (
            <Select
              label="Mazdoor"
              name="mazdoor"
              register={register}
              options={mazdoorOptions}
              error={errors.mazdoor?.message}
            />
          )}
          {category === 'raw_material' && (
            <Select
              label="Supplier"
              name="supplier"
              register={register}
              options={supplierOptions}
              error={errors.supplier?.message}
            />
          )}
          <Input
            label="Bill Number"
            name="billNumber"
            register={register}
            error={errors.billNumber?.message}
            placeholder="Optional"
          />
          <div className="mt-4">
            <label className="label">Notes</label>
            <textarea
              {...register('notes')}
              rows={3}
              className="input"
              placeholder="Additional notes..."
            />
          </div>

          <div className="flex justify-end gap-3 mt-6">
            <Button
              type="button"
              variant="secondary"
              onClick={() => {
                setIsModalOpen(false)
                reset()
              }}
            >
              Cancel
            </Button>
            <Button type="submit" variant="primary">
              {editingExpense ? 'Update' : 'Create'} Expense
            </Button>
          </div>
        </form>
      </Modal>
    </div>
  )
}

export default Expenses

