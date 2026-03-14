import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Search, Edit2, Trash2 } from 'lucide-react'
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

const LabourExpenses = () => {
  const [labourExpenses, setLabourExpenses] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingLabourExpense, setEditingLabourExpense] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')

  const { register, handleSubmit, formState: { errors }, reset, setValue } = useForm()

  useEffect(() => {
    fetchLabourExpenses()
  }, [searchTerm])

  const fetchLabourExpenses = async () => {
    try {
      setLoading(true)
      const response = await api.get('/labour-expenses', {
        params: {
          page: 1,
          limit: 1000,
          search: searchTerm,
          isActive: true,
        },
      })
      setLabourExpenses(response.data.data || [])
    } catch (error) {
      toast.error('Failed to fetch labour expenses')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleAddLabourExpense = () => {
    setEditingLabourExpense(null)
    reset({
      name: '',
      rate: ''
    })
    setIsModalOpen(true)
  }

  const handleEditLabourExpense = (labourExpense) => {
    setEditingLabourExpense(labourExpense)
    setValue('name', labourExpense.name)
    setValue('rate', labourExpense.rate)
    setIsModalOpen(true)
  }

  const handleDeleteLabourExpense = async (row) => {
    if (!window.confirm('Are you sure you want to delete this labour expense?')) {
      return
    }

    const id = row?._id || row?.id
    if (!id || typeof id !== 'string') {
      toast.error('Invalid labour expense')
      return
    }

    try {
      await api.delete(`/labour-expenses/${id}`)
      toast.success('Labour expense deleted successfully')
      // Optimistically update table
      setLabourExpenses(prev => prev.filter(le => (le._id || le.id) !== id))
      fetchLabourExpenses()
    } catch (error) {
      toast.error('Failed to delete labour expense')
      console.error(error)
    }
  }

  const onSubmit = async (data) => {
    try {
      if (editingLabourExpense) {
        await api.put(`/labour-expenses/${editingLabourExpense._id}`, data)
        toast.success('Labour expense updated successfully')
      } else {
        await api.post('/labour-expenses', data)
        toast.success('Labour expense created successfully')
      }
      setIsModalOpen(false)
      fetchLabourExpenses()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save labour expense')
      console.error(error)
    }
  }

  const columns = [
    {
      key: 'name',
      label: 'Name',
      render: (value, row) => (
        <div className="font-medium text-gray-900">{value}</div>
      ),
    },
    {
      key: 'rate',
      label: 'Rate',
      render: (value, row) => (
        <div className="flex items-center">
          <span className="text-sm font-bold text-green-600 mr-1">PKR</span>
          <span className="font-medium">{parseFloat(value).toFixed(2)}</span>
        </div>
      ),
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
          <h1 className="text-2xl font-bold text-gray-900">Labour Expenses</h1>
          <p className="text-gray-600">Manage labour expenses rates for your business</p>
        </div>
        <Button onClick={handleAddLabourExpense}>
          <Plus className="h-4 w-4 mr-2" />
          Add Labour Expense
        </Button>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>Labour Expenses</CardTitle>
          <CardDescription>
            View and manage all labour expense rates in your system
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="flex items-center gap-4 mb-6">
            <div className="relative flex-1 max-w-sm">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-4 w-4" />
              <input
                type="text"
                placeholder="Search labour expenses..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent"
              />
            </div>
          </div>

          <DataTable
            columns={columns}
            data={labourExpenses}
            loading={loading}
            onEdit={handleEditLabourExpense}
            onDelete={handleDeleteLabourExpense}
            searchPlaceholder="Search labour expenses..."
          />
        </CardContent>
      </Card>

      <Dialog open={isModalOpen} onOpenChange={setIsModalOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>
              {editingLabourExpense ? 'Edit Labour Expense' : 'Add Labour Expense'}
            </DialogTitle>
            <DialogDescription>
              {editingLabourExpense
                ? 'Update the labour expense information.'
                : 'Add a new labour expense to your system.'}
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(onSubmit)}>
            <div className="space-y-4">
              <FormInput
                label="Name"
                name="name"
                register={register}
                errors={errors}
                placeholder="Enter labour expense name"
                required
              />
              <FormInput
                label="Rate"
                name="rate"
                type="number"
                step="0.01"
                register={register}
                errors={errors}
                placeholder="Enter rate"
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
              <Button type="submit">
                {editingLabourExpense ? 'Update' : 'Create'}
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}

export default LabourExpenses
