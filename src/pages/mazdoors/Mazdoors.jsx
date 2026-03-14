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
import { formatCurrency, formatDate, formatPhone } from '../../utils/formatters'

const Mazdoors = () => {
  const [mazdoors, setMazdoors] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingMazdoor, setEditingMazdoor] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [pagination, setPagination] = useState({ page: 1, limit: 10, total: 0 })

  const { register, handleSubmit, formState: { errors }, reset, setValue } = useForm()

  useEffect(() => {
    fetchMazdoors()
  }, [pagination.page, searchTerm])

  const fetchMazdoors = async () => {
    try {
      setLoading(true)
      const response = await api.get('/mazdoors', {
        params: {
          page: pagination.page,
          limit: pagination.limit,
          search: searchTerm,
        },
      })
      setMazdoors(response.data.data || [])
      setPagination(prev => ({
        ...prev,
        total: response.data.pagination?.total || 0,
      }))
    } catch (error) {
      toast.error('Failed to fetch mazdoors')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    setEditingMazdoor(null)
    reset()
    setValue('hireDate', new Date().toISOString().split('T')[0])
    setValue('salaryType', 'daily')
    setIsModalOpen(true)
  }

  const handleEdit = (mazdoor) => {
    setEditingMazdoor(mazdoor)
    setValue('name', mazdoor.name)
    setValue('phone', mazdoor.phone)
    setValue('alternatePhone', mazdoor.alternatePhone || '')
    setValue('cnic', mazdoor.cnic || '')
    setValue('address.street', mazdoor.address?.street || '')
    setValue('address.city', mazdoor.address?.city || '')
    setValue('address.state', mazdoor.address?.state || '')
    setValue('address.zipCode', mazdoor.address?.zipCode || '')
    setValue('hireDate', mazdoor.hireDate ? mazdoor.hireDate.split('T')[0] : new Date().toISOString().split('T')[0])
    setValue('salary', mazdoor.salary || 0)
    setValue('salaryType', mazdoor.salaryType || 'daily')
    setValue('notes', mazdoor.notes || '')
    setIsModalOpen(true)
  }

  const handleDelete = async (mazdoor) => {
    if (!window.confirm(`Are you sure you want to delete ${mazdoor.name}?`)) {
      return
    }

    try {
      await api.delete(`/mazdoors/${mazdoor._id}`)
      toast.success('Mazdoor deleted successfully')
      fetchMazdoors()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete mazdoor')
    }
  }

  const onSubmit = async (data) => {
    try {
      if (editingMazdoor) {
        await api.put(`/mazdoors/${editingMazdoor._id}`, data)
        toast.success('Mazdoor updated successfully')
      } else {
        await api.post('/mazdoors', data)
        toast.success('Mazdoor created successfully')
      }
      setIsModalOpen(false)
      reset()
      fetchMazdoors()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save mazdoor')
    }
  }

  const columns = [
    { key: 'name', label: 'Name' },
    { key: 'phone', label: 'Phone', render: (value) => formatPhone(value) },
    { key: 'cnic', label: 'CNIC', render: (value) => value || '-' },
    {
      key: 'salary',
      label: 'Salary',
      render: (value, row) => `${formatCurrency(value || 0)} / ${row.salaryType || 'daily'}`,
    },
    {
      key: 'currentBalance',
      label: 'Balance',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'hireDate',
      label: 'Hire Date',
      render: (value) => formatDate(value),
    },
    {
      key: 'isActive',
      label: 'Status',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs ${
          value ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'
        }`}>
          {value ? 'Active' : 'Inactive'}
        </span>
      ),
    },
  ]

  const actions = [
    { key: 'edit', label: 'Edit', variant: 'primary' },
    { key: 'delete', label: 'Delete', variant: 'danger' },
  ]

  const handleAction = (action, mazdoor) => {
    if (action === 'edit') {
      handleEdit(mazdoor)
    } else if (action === 'delete') {
      handleDelete(mazdoor)
    }
  }

  const salaryTypeOptions = [
    { value: 'daily', label: 'Daily' },
    { value: 'weekly', label: 'Weekly' },
    { value: 'monthly', label: 'Monthly' },
    { value: 'per_work', label: 'Per Work' },
  ]

  return (
    <div>
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Mazdoors</h1>
          <p className="text-gray-600 mt-1">Manage workers (mazdoors)</p>
        </div>
        <Button onClick={handleCreate} variant="primary">
          <Plus className="h-5 w-5 mr-2" />
          Add Mazdoor
        </Button>
      </div>

      <Card>
        <div className="mb-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search mazdoors..."
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
            data={mazdoors}
            actions={actions}
            onAction={handleAction}
          />
        )}

        {pagination.total > pagination.limit && (
          <div className="mt-4 flex justify-between items-center">
            <p className="text-sm text-gray-600">
              Showing {((pagination.page - 1) * pagination.limit) + 1} to{' '}
              {Math.min(pagination.page * pagination.limit, pagination.total)} of{' '}
              {pagination.total} mazdoors
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
        title={editingMazdoor ? 'Edit Mazdoor' : 'Add New Mazdoor'}
        size="lg"
      >
        <form onSubmit={handleSubmit(onSubmit)}>
          <div className="grid grid-cols-2 gap-4">
            <Input
              label="Name"
              name="name"
              register={register}
              required
              error={errors.name?.message}
              placeholder="Mazdoor name"
            />
            <Input
              label="Phone"
              name="phone"
              register={register}
              required
              error={errors.phone?.message}
              placeholder="03XX-XXXXXXX"
            />
            <Input
              label="Alternate Phone"
              name="alternatePhone"
              register={register}
              error={errors.alternatePhone?.message}
              placeholder="Optional"
            />
            <Input
              label="CNIC"
              name="cnic"
              register={register}
              error={errors.cnic?.message}
              placeholder="XXXXX-XXXXXXX-X"
            />
            <Input
              label="Street Address"
              name="address.street"
              register={register}
              error={errors.address?.street?.message}
              placeholder="Street address"
            />
            <Input
              label="City"
              name="address.city"
              register={register}
              error={errors.address?.city?.message}
              placeholder="City"
            />
            <Input
              label="State"
              name="address.state"
              register={register}
              error={errors.address?.state?.message}
              placeholder="State/Province"
            />
            <Input
              label="Zip Code"
              name="address.zipCode"
              register={register}
              error={errors.address?.zipCode?.message}
              placeholder="Zip code"
            />
            <Input
              label="Hire Date"
              name="hireDate"
              type="date"
              register={register}
              error={errors.hireDate?.message}
            />
            <Select
              label="Salary Type"
              name="salaryType"
              register={register}
              options={salaryTypeOptions}
              error={errors.salaryType?.message}
            />
            <Input
              label="Salary"
              name="salary"
              type="number"
              register={register}
              error={errors.salary?.message}
              placeholder="0"
            />
          </div>
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
              {editingMazdoor ? 'Update' : 'Create'} Mazdoor
            </Button>
          </div>
        </form>
      </Modal>
    </div>
  )
}

export default Mazdoors

