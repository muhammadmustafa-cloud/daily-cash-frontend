import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Edit, Trash2, Search, AlertCircle } from 'lucide-react'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Select from '../../components/ui/Select'
import Modal from '../../components/ui/Modal'
import Table from '../../components/ui/Table'
import api from '../../config/api'
import { formatCurrency, formatDate } from '../../utils/formatters'

const Items = () => {
  const [items, setItems] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingItem, setEditingItem] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [pagination, setPagination] = useState({ page: 1, limit: 10, total: 0 })

  const { register, handleSubmit, formState: { errors }, reset, setValue, watch } = useForm()
  const itemType = watch('type')

  useEffect(() => {
    fetchItems()
  }, [pagination.page, searchTerm])

  const fetchItems = async () => {
    try {
      setLoading(true)
      const response = await api.get('/items', {
        params: {
          page: pagination.page,
          limit: pagination.limit,
          search: searchTerm,
        },
      })
      setItems(response.data.data || [])
      setPagination(prev => ({
        ...prev,
        total: response.data.pagination?.total || 0,
      }))
    } catch (error) {
      toast.error('Failed to fetch items')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    setEditingItem(null)
    reset()
    setIsModalOpen(true)
  }

  const handleEdit = (item) => {
    setEditingItem(item)
    setValue('name', item.name)
    setValue('code', item.code)
    setValue('type', item.type)
    setValue('category', item.category || '')
    setValue('unit', item.unit || 'kg')
    setValue('description', item.description || '')
    setValue('purchasePrice', item.purchasePrice || 0)
    setValue('sellingPrice', item.sellingPrice || 0)
    setValue('minStockLevel', item.minStockLevel || 0)
    setValue('maxStockLevel', item.maxStockLevel || 0)
    setValue('reorderPoint', item.reorderPoint || 0)
    setValue('conversionRate', item.conversionRate || 1)
    setValue('notes', item.notes || '')
    setIsModalOpen(true)
  }

  const handleDelete = async (item) => {
    if (!window.confirm(`Are you sure you want to delete ${item.name}?`)) {
      return
    }

    try {
      await api.delete(`/items/${item._id}`)
      toast.success('Item deleted successfully')
      fetchItems()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete item')
    }
  }

  const onSubmit = async (data) => {
    try {
      if (editingItem) {
        await api.put(`/items/${editingItem._id}`, data)
        toast.success('Item updated successfully')
      } else {
        await api.post('/items', data)
        toast.success('Item created successfully')
      }
      setIsModalOpen(false)
      reset()
      fetchItems()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save item')
    }
  }

  const columns = [
    { key: 'code', label: 'Code' },
    { key: 'name', label: 'Name' },
    { 
      key: 'type', 
      label: 'Type',
      render: (value) => value?.replace('_', ' ').toUpperCase() || '-'
    },
    { key: 'category', label: 'Category', render: (value) => value || '-' },
    { key: 'unit', label: 'Unit' },
    { 
      key: 'currentStock', 
      label: 'Stock',
      render: (value, row) => (
        <div className="flex items-center gap-2">
          <span>{value || 0}</span>
          {row.minStockLevel > 0 && value <= row.minStockLevel && (
            <AlertCircle className="h-4 w-4 text-orange-500" title="Low stock" />
          )}
        </div>
      )
    },
    { 
      key: 'sellingPrice', 
      label: 'Selling Price',
      render: (value) => formatCurrency(value || 0)
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

  const handleAction = (action, item) => {
    if (action === 'edit') {
      handleEdit(item)
    } else if (action === 'delete') {
      handleDelete(item)
    }
  }

  const itemTypeOptions = [
    { value: 'raw_material', label: 'Raw Material' },
    { value: 'finished_product', label: 'Finished Product' },
    { value: 'packaging', label: 'Packaging' },
    { value: 'other', label: 'Other' },
  ]

  const unitOptions = [
    { value: 'kg', label: 'Kilogram (kg)' },
    { value: 'g', label: 'Gram (g)' },
    { value: 'ton', label: 'Ton' },
    { value: 'quintal', label: 'Quintal' },
    { value: 'bag', label: 'Bag' },
    { value: 'piece', label: 'Piece' },
    { value: 'liter', label: 'Liter' },
    { value: 'meter', label: 'Meter' },
  ]

  return (
    <div>
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Items</h1>
          <p className="text-gray-600 mt-1">Manage your inventory items</p>
        </div>
        <Button onClick={handleCreate} variant="primary">
          <Plus className="h-5 w-5 mr-2" />
          Add Item
        </Button>
      </div>

      <Card>
        <div className="mb-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search items..."
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
            data={items}
            actions={actions}
            onAction={handleAction}
          />
        )}

        {pagination.total > pagination.limit && (
          <div className="mt-4 flex justify-between items-center">
            <p className="text-sm text-gray-600">
              Showing {((pagination.page - 1) * pagination.limit) + 1} to{' '}
              {Math.min(pagination.page * pagination.limit, pagination.total)} of{' '}
              {pagination.total} items
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
        title={editingItem ? 'Edit Item' : 'Add New Item'}
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
              placeholder="Item name"
            />
            {editingItem && (
              <Input
                label="Code"
                name="code"
                register={register}
                error={errors.code?.message}
                placeholder="Item code"
                disabled
              />
            )}
            <Select
              label="Type"
              name="type"
              register={register}
              required
              options={itemTypeOptions}
              error={errors.type?.message}
            />
            <Input
              label="Category"
              name="category"
              register={register}
              error={errors.category?.message}
              placeholder="Item category"
            />
            <Select
              label="Unit"
              name="unit"
              register={register}
              required
              options={unitOptions}
              error={errors.unit?.message}
            />
            <Input
              label="Purchase Price"
              name="purchasePrice"
              type="number"
              step="0.01"
              register={register}
              error={errors.purchasePrice?.message}
              placeholder="0.00"
            />
            <Input
              label="Selling Price"
              name="sellingPrice"
              type="number"
              step="0.01"
              register={register}
              error={errors.sellingPrice?.message}
              placeholder="0.00"
            />
            <Input
              label="Min Stock Level"
              name="minStockLevel"
              type="number"
              register={register}
              error={errors.minStockLevel?.message}
              placeholder="0"
            />
            <Input
              label="Max Stock Level"
              name="maxStockLevel"
              type="number"
              register={register}
              error={errors.maxStockLevel?.message}
              placeholder="0"
            />
            <Input
              label="Reorder Point"
              name="reorderPoint"
              type="number"
              register={register}
              error={errors.reorderPoint?.message}
              placeholder="0"
            />
            {itemType === 'finished_product' && (
              <Input
                label="Conversion Rate"
                name="conversionRate"
                type="number"
                step="0.01"
                register={register}
                error={errors.conversionRate?.message}
                placeholder="1.0"
              />
            )}
          </div>
          <div className="mt-4">
            <label className="label">Description</label>
            <textarea
              {...register('description')}
              rows={3}
              className="input"
              placeholder="Item description..."
            />
          </div>
          <div className="mt-4">
            <label className="label">Notes</label>
            <textarea
              {...register('notes')}
              rows={2}
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
              {editingItem ? 'Update' : 'Create'} Item
            </Button>
          </div>
        </form>
      </Modal>
    </div>
  )
}

export default Items

