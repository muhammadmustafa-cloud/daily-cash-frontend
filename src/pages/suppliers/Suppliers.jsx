import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Edit, Trash2, Search } from 'lucide-react'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Modal from '../../components/ui/Modal'
import Table from '../../components/ui/Table'
import api from '../../config/api'
import { formatDate, formatPhone } from '../../utils/formatters'

const Suppliers = () => {
  const [suppliers, setSuppliers] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingSupplier, setEditingSupplier] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [pagination, setPagination] = useState({ page: 1, limit: 10, total: 0 })

  const { register, handleSubmit, formState: { errors }, reset, setValue } = useForm()

  useEffect(() => {
    fetchSuppliers()
  }, [pagination.page, searchTerm])

  const fetchSuppliers = async () => {
    try {
      setLoading(true)
      const response = await api.get('/suppliers', {
        params: {
          page: pagination.page,
          limit: pagination.limit,
          search: searchTerm,
        },
      })
      setSuppliers(response.data.data || [])
      setPagination(prev => ({
        ...prev,
        total: response.data.pagination?.total || 0,
      }))
    } catch (error) {
      toast.error('Failed to fetch suppliers')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    setEditingSupplier(null)
    reset()
    setIsModalOpen(true)
  }

  const handleEdit = (supplier) => {
    setEditingSupplier(supplier)
    setValue('name', supplier.name)
    setValue('address', supplier.address || '')
    setValue('phone', supplier.phone || '')
    setIsModalOpen(true)
  }

  const handleDelete = async (supplier) => {
    if (!window.confirm(`Are you sure you want to delete ${supplier.name}?`)) {
      return
    }

    try {
      await api.delete(`/suppliers/${supplier._id}`)
      toast.success('Supplier deleted successfully')
      fetchSuppliers()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete supplier')
    }
  }

  const onSubmit = async (data) => {
    try {
      if (editingSupplier) {
        await api.put(`/suppliers/${editingSupplier._id}`, data)
        toast.success('Supplier updated successfully')
      } else {
        await api.post('/suppliers', data)
        toast.success('Supplier created successfully')
      }
      setIsModalOpen(false)
      reset()
      fetchSuppliers()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save supplier')
    }
  }

  const columns = [
    { key: 'name', label: 'Name' },
    { key: 'address', label: 'Address' },
    { key: 'phone', label: 'Contact No', render: (value) => value ? formatPhone(value) : '-' },
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
    {
      key: 'createdAt',
      label: 'Created',
      render: (value) => formatDate(value),
    },
  ]

  const actions = [
    { key: 'edit', label: 'Edit', variant: 'primary' },
    { key: 'delete', label: 'Delete', variant: 'danger' },
  ]

  const handleAction = (action, supplier) => {
    if (action === 'edit') {
      handleEdit(supplier)
    } else if (action === 'delete') {
      handleDelete(supplier)
    }
  }

  return (
    <div>
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Suppliers</h1>
          <p className="text-gray-600 mt-1">Manage your suppliers</p>
        </div>
        <Button onClick={handleCreate} variant="primary">
          <Plus className="h-5 w-5 mr-2" />
          Add Supplier
        </Button>
      </div>

      <Card>
        <div className="mb-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search suppliers..."
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
            data={suppliers}
            actions={actions}
            onAction={handleAction}
          />
        )}

        {pagination.total > pagination.limit && (
          <div className="mt-4 flex justify-between items-center">
            <p className="text-sm text-gray-600">
              Showing {((pagination.page - 1) * pagination.limit) + 1} to{' '}
              {Math.min(pagination.page * pagination.limit, pagination.total)} of{' '}
              {pagination.total} suppliers
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
        title={editingSupplier ? 'Edit Supplier' : 'Add New Supplier'}
        size="lg"
      >
        <form onSubmit={handleSubmit(onSubmit)}>
          <Input
            label="Name"
            name="name"
            register={register}
            required
            error={errors.name?.message}
            placeholder="Supplier name"
          />
          <Input
            label="Address"
            name="address"
            register={register}
            required
            error={errors.address?.message}
            placeholder="Supplier address"
          />
          <Input
            label="Contact No"
            name="phone"
            register={register}
            error={errors.phone?.message}
            placeholder="03XX-XXXXXXX (Optional)"
          />

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
              {editingSupplier ? 'Update' : 'Create'} Supplier
            </Button>
          </div>
        </form>
      </Modal>
    </div>
  )
}

export default Suppliers

