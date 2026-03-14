import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Search, TrendingUp, TrendingDown } from 'lucide-react'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Select from '../../components/ui/Select'
import Modal from '../../components/ui/Modal'
import Table from '../../components/ui/Table'
import api from '../../config/api'
import { formatCurrency, formatDate } from '../../utils/formatters'

const Inventory = () => {
  const [stockRecords, setStockRecords] = useState([])
  const [items, setItems] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [searchTerm, setSearchTerm] = useState('')
  const [pagination, setPagination] = useState({ page: 1, limit: 10, total: 0 })

  const { register, handleSubmit, formState: { errors }, reset, watch, setValue } = useForm()
  const operation = watch('operation')

  useEffect(() => {
    fetchStockRecords()
    fetchItems()
  }, [pagination.page, searchTerm])

  const fetchItems = async () => {
    try {
      const response = await api.get('/items?limit=1000')
      setItems(response.data.data || [])
    } catch (error) {
      console.error('Error fetching items:', error)
    }
  }

  const fetchStockRecords = async () => {
    try {
      setLoading(true)
      const response = await api.get('/inventory', {
        params: {
          page: pagination.page,
          limit: pagination.limit,
          search: searchTerm,
        },
      })
      setStockRecords(response.data.data || [])
      setPagination(prev => ({
        ...prev,
        total: response.data.pagination?.total || 0,
      }))
    } catch (error) {
      toast.error('Failed to fetch inventory records')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    reset()
    setValue('operation', 'in')
    setValue('date', new Date().toISOString().split('T')[0])
    setIsModalOpen(true)
  }

  const onSubmit = async (data) => {
    try {
      await api.post('/inventory', data)
      toast.success('Stock record created successfully')
      setIsModalOpen(false)
      reset()
      fetchStockRecords()
      fetchItems()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to create stock record')
    }
  }

  const columns = [
    {
      key: 'item',
      label: 'Item',
      render: (value) => (value?.name || '-'),
    },
    {
      key: 'operation',
      label: 'Operation',
      render: (value) => {
        const ops = {
          in: { label: 'Stock In', color: 'bg-green-100 text-green-800' },
          out: { label: 'Stock Out', color: 'bg-red-100 text-red-800' },
          adjustment: { label: 'Adjustment', color: 'bg-yellow-100 text-yellow-800' },
          transfer: { label: 'Transfer', color: 'bg-blue-100 text-blue-800' },
        }
        const op = ops[value] || ops.in
        return <span className={`px-2 py-1 rounded-full text-xs ${op.color}`}>{op.label}</span>
      },
    },
    { key: 'quantity', label: 'Quantity' },
    {
      key: 'rate',
      label: 'Rate',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'totalAmount',
      label: 'Total',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'previousStock',
      label: 'Previous Stock',
    },
    {
      key: 'newStock',
      label: 'New Stock',
    },
    {
      key: 'date',
      label: 'Date',
      render: (value) => formatDate(value),
    },
  ]

  const itemOptions = items.map(item => ({
    value: item._id,
    label: `${item.name} (${item.code || 'N/A'})`,
  }))

  const operationOptions = [
    { value: 'in', label: 'Stock In' },
    { value: 'out', label: 'Stock Out' },
    { value: 'adjustment', label: 'Adjustment' },
    { value: 'transfer', label: 'Transfer' },
  ]

  return (
    <div>
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Inventory</h1>
          <p className="text-gray-600 mt-1">Manage stock in/out and adjustments</p>
        </div>
        <Button onClick={handleCreate} variant="primary">
          <Plus className="h-5 w-5 mr-2" />
          Stock Entry
        </Button>
      </div>

      <Card>
        <div className="mb-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search inventory records..."
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
          <Table columns={columns} data={stockRecords} />
        )}

        {pagination.total > pagination.limit && (
          <div className="mt-4 flex justify-between items-center">
            <p className="text-sm text-gray-600">
              Showing {((pagination.page - 1) * pagination.limit) + 1} to{' '}
              {Math.min(pagination.page * pagination.limit, pagination.total)} of{' '}
              {pagination.total} records
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
        title="Stock Entry"
        size="md"
      >
        <form onSubmit={handleSubmit(onSubmit)}>
          <Select
            label="Item"
            name="item"
            register={register}
            required
            options={itemOptions}
            error={errors.item?.message}
            placeholder="Select item"
          />
          <Select
            label="Operation"
            name="operation"
            register={register}
            required
            options={operationOptions}
            error={errors.operation?.message}
          />
          <div className="grid grid-cols-2 gap-4">
            <Input
              label="Quantity"
              name="quantity"
              type="number"
              step="0.01"
              register={register}
              required
              error={errors.quantity?.message}
              placeholder="0.00"
            />
            <Input
              label="Rate"
              name="rate"
              type="number"
              step="0.01"
              register={register}
              error={errors.rate?.message}
              placeholder="0.00"
            />
          </div>
          <Input
            label="Date"
            name="date"
            type="date"
            register={register}
            required
            error={errors.date?.message}
          />
          {operation === 'adjustment' && (
            <Input
              label="Reason"
              name="reason"
              register={register}
              error={errors.reason?.message}
              placeholder="Reason for adjustment"
            />
          )}
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
              Create Entry
            </Button>
          </div>
        </form>
      </Modal>
    </div>
  )
}

export default Inventory

