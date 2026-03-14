import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Search, Filter, X, Plus, ShoppingCart, Package, Eye, Trash2 } from 'lucide-react'
import Card, { CardContent, CardDescription, CardHeader, CardTitle } from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import { FormInput } from '../../components/ui/form-input'
import { FormSelect } from '../../components/ui/form-select'
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
import { formatCurrency, formatDate } from '../../utils/formatters'

const Transactions = () => {
  const [transactions, setTransactions] = useState([])
  const [customers, setCustomers] = useState([])
  const [suppliers, setSuppliers] = useState([])
  const [items, setItems] = useState([])
  const [accounts, setAccounts] = useState([])
  const [loading, setLoading] = useState(false)
  const [searchTerm, setSearchTerm] = useState('')
  const [pagination, setPagination] = useState({ page: 1, limit: 10, total: 0 })
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [transactionType, setTransactionType] = useState('sale')
  const [viewingTransaction, setViewingTransaction] = useState(null)
  const [isViewModalOpen, setIsViewModalOpen] = useState(false)
  
  // Filter states
  const [filters, setFilters] = useState({
    type: '',
    customer: '',
    supplier: '',
    paymentStatus: '',
    paymentMethod: '',
    startDate: '',
    endDate: '',
  })
  const [showFilters, setShowFilters] = useState(false)

  // Simplified item selection
  const [selectedItem, setSelectedItem] = useState(null)
  const [itemQuantity, setItemQuantity] = useState(1)
  const [transactionItems, setTransactionItems] = useState([])

  const { register, handleSubmit, formState: { errors }, reset, setValue, watch } = useForm()
  const discount = watch('discount') || 0
  const tax = watch('tax') || 0
  const paidAmount = watch('paidAmount') || 0
  const account = watch('account')

  useEffect(() => {
    fetchTransactions()
    fetchCustomers()
    fetchSuppliers()
    fetchItems()
    fetchAccounts()
  }, [pagination.page, searchTerm, filters])

  const fetchTransactions = async () => {
    try {
      setLoading(true)
      const params = {
        page: pagination.page,
        limit: pagination.limit,
        search: searchTerm,
      }

      if (filters.type) params.type = filters.type
      if (filters.customer) params.customer = filters.customer
      if (filters.supplier) params.supplier = filters.supplier
      if (filters.paymentStatus) params.paymentStatus = filters.paymentStatus
      if (filters.paymentMethod) params.paymentMethod = filters.paymentMethod
      if (filters.startDate) params.startDate = filters.startDate
      if (filters.endDate) params.endDate = filters.endDate

      const response = await api.get('/transactions', { params })
      setTransactions(response.data.data || [])
      setPagination(prev => ({
        ...prev,
        total: response.data.pagination?.total || 0,
      }))
    } catch (error) {
      toast.error('Failed to fetch transactions')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const fetchCustomers = async () => {
    try {
      const response = await api.get('/customers?limit=1000')
      setCustomers(response.data.data || [])
    } catch (error) {
      console.error('Error fetching customers:', error)
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

  const fetchItems = async () => {
    try {
      const response = await api.get('/items?limit=1000&isActive=true')
      setItems(response.data.data || [])
    } catch (error) {
      console.error('Error fetching items:', error)
    }
  }

  const fetchAccounts = async () => {
    try {
      const response = await api.get('/accounts?limit=1000&isActive=true')
      setAccounts(response.data.data || [])
    } catch (error) {
      console.error('Error fetching accounts:', error)
    }
  }

  const handleCreate = (type) => {
    setTransactionType(type)
    setTransactionItems([])
    setSelectedItem(null)
    setItemQuantity(1)
    reset({
      date: new Date().toISOString().split('T')[0],
      discount: 0,
      tax: 0,
      paidAmount: 0,
      paymentMethod: 'cash',
    })
    setIsModalOpen(true)
  }

  const handleView = async (transaction) => {
    try {
      const response = await api.get(`/transactions/${transaction._id}`)
      setViewingTransaction(response.data.transaction)
      setIsViewModalOpen(true)
    } catch (error) {
      toast.error('Failed to fetch transaction details')
    }
  }

  const handleDelete = async (transaction) => {
    if (!window.confirm('Are you sure you want to delete this transaction?')) {
      return
    }

    try {
      await api.delete(`/transactions/${transaction._id}`)
      toast.success('Transaction deleted successfully')
      fetchTransactions()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete transaction')
    }
  }

  const addItemToTransaction = () => {
    if (!selectedItem) {
      toast.error('Please select an item')
      return
    }

    if (transactionType === 'sale' && selectedItem.currentStock < itemQuantity) {
      toast.error(`Insufficient stock. Available: ${selectedItem.currentStock} ${selectedItem.unit}`)
      return
    }

    const rate = transactionType === 'sale' ? selectedItem.sellingPrice : selectedItem.purchasePrice
    const existingIndex = transactionItems.findIndex(ti => ti.item._id === selectedItem._id)

    if (existingIndex >= 0) {
      const updated = [...transactionItems]
      updated[existingIndex].quantity += parseFloat(itemQuantity) || 0
      updated[existingIndex].total = updated[existingIndex].quantity * updated[existingIndex].rate
      setTransactionItems(updated)
    } else {
      setTransactionItems([
        ...transactionItems,
        {
          item: selectedItem,
          quantity: parseFloat(itemQuantity) || 1,
          rate: rate || 0,
          total: (parseFloat(itemQuantity) || 1) * (rate || 0),
        }
      ])
    }

    setSelectedItem(null)
    setItemQuantity(1)
  }

  const removeItem = (index) => {
    setTransactionItems(transactionItems.filter((_, i) => i !== index))
  }

  const updateItemQuantity = (index, quantity) => {
    const updated = [...transactionItems]
    const qty = parseFloat(quantity) || 0
    
    if (transactionType === 'sale' && qty > updated[index].item.currentStock) {
      toast.error(`Insufficient stock. Available: ${updated[index].item.currentStock} ${updated[index].item.unit}`)
      return
    }
    
    updated[index].quantity = qty
    updated[index].total = qty * updated[index].rate
    setTransactionItems(updated)
  }

  const calculateTotals = () => {
    const subtotal = transactionItems.reduce((sum, item) => sum + item.total, 0)
    const finalDiscount = parseFloat(discount) || 0
    const finalTax = parseFloat(tax) || 0
    const total = subtotal - finalDiscount + finalTax
    const finalPaidAmount = parseFloat(paidAmount) || 0
    const remaining = total - finalPaidAmount

    return { subtotal, discount: finalDiscount, tax: finalTax, total, paidAmount: finalPaidAmount, remaining }
  }

  const onSubmit = async (data) => {
    try {
      if (transactionItems.length === 0) {
        toast.error('Please add at least one item')
        return
      }

      const totals = calculateTotals()

      const transactionData = {
        type: transactionType,
        [transactionType === 'sale' ? 'customer' : 'supplier']: data[transactionType === 'sale' ? 'customer' : 'supplier'],
        items: transactionItems.map(ti => ({
          item: ti.item._id,
          quantity: ti.quantity,
          rate: ti.rate,
        })),
        subtotal: totals.subtotal,
        discount: totals.discount,
        tax: totals.tax,
        total: totals.total,
        paidAmount: totals.paidAmount,
        paymentMethod: data.paymentMethod,
        date: data.date,
        notes: data.notes || '',
      }

      await api.post('/transactions', transactionData)

      // If payment is made, create a payment/receipt record
      if (totals.paidAmount > 0 && account) {
        try {
          const customerOrSupplier = transactionType === 'sale' 
            ? customers.find(c => c._id === data.customer)
            : suppliers.find(s => s._id === data.supplier)
          
          const paymentMethodMap = {
            'cash': 'cash',
            'bank': 'bank_transfer',
            'cheque': 'cheque',
            'online': 'online'
          }

          const paymentData = {
            type: transactionType === 'sale' ? 'receipt' : 'payment',
            date: data.date,
            description: `${transactionType === 'sale' ? 'Sale' : 'Purchase'} payment from ${customerOrSupplier?.name || 'N/A'}`,
            amount: totals.paidAmount,
            paymentMethod: paymentMethodMap[data.paymentMethod] || 'cash',
            [transactionType === 'sale' ? 'toAccount' : 'fromAccount']: account,
            [transactionType === 'sale' ? 'customer' : 'supplier']: data[transactionType === 'sale' ? 'customer' : 'supplier'],
            category: transactionType === 'sale' ? 'customer_payment' : 'supplier_payment',
          }
          await api.post('/payments', paymentData)
        } catch (error) {
          console.error('Error creating payment record:', error)
          // Don't fail the transaction if payment record fails
        }
      }

      toast.success(`${transactionType === 'sale' ? 'Sale' : 'Purchase'} created successfully`)
      setIsModalOpen(false)
      reset()
      setTransactionItems([])
      setSelectedItem(null)
      fetchTransactions()
    } catch (error) {
      toast.error(error.response?.data?.message || `Failed to create ${transactionType}`)
    }
  }

  const handleFilterChange = (key, value) => {
    setFilters(prev => ({
      ...prev,
      [key]: value,
    }))
    setPagination(prev => ({ ...prev, page: 1 }))
  }

  const clearFilters = () => {
    setFilters({
      type: '',
      customer: '',
      supplier: '',
      paymentStatus: '',
      paymentMethod: '',
      startDate: '',
      endDate: '',
    })
    setSearchTerm('')
    setPagination(prev => ({ ...prev, page: 1 }))
  }

  const activeFiltersCount = Object.values(filters).filter(v => v !== '').length + (searchTerm ? 1 : 0)
  const totals = calculateTotals()

  const columns = [
    {
      key: 'transactionNumber',
      label: 'Transaction #',
    },
    {
      key: 'type',
      label: 'Type',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs font-semibold ${
          value === 'sale' ? 'bg-green-100 text-green-800' :
          value === 'purchase' ? 'bg-blue-100 text-blue-800' :
          'bg-gray-100 text-gray-800'
        }`}>
          {value?.toUpperCase() || '-'}
        </span>
      ),
    },
    {
      key: 'customer',
      label: 'Customer',
      render: (value) => value?.name || '-',
    },
    {
      key: 'supplier',
      label: 'Supplier',
      render: (value) => value?.name || '-',
    },
    {
      key: 'items',
      label: 'Items',
      render: (value) => value?.length || 0,
    },
    {
      key: 'total',
      label: 'Total',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'paidAmount',
      label: 'Paid',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'remainingAmount',
      label: 'Remaining',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'paymentStatus',
      label: 'Status',
      render: (value) => {
        const statuses = {
          pending: { label: 'Pending', color: 'bg-yellow-100 text-yellow-800' },
          partial: { label: 'Partial', color: 'bg-orange-100 text-orange-800' },
          paid: { label: 'Paid', color: 'bg-green-100 text-green-800' },
          overdue: { label: 'Overdue', color: 'bg-red-100 text-red-800' },
        }
        const status = statuses[value] || statuses.pending
        return <span className={`px-2 py-1 rounded-full text-xs font-semibold ${status.color}`}>{status.label}</span>
      },
    },
    {
      key: 'date',
      label: 'Date',
      render: (value) => formatDate(value),
    },
  ]

  const customerOptions = customers.map(c => ({
    value: c._id,
    label: `${c.name}${c.phone ? ` - ${c.phone}` : ''}`
  }))

  const supplierOptions = suppliers.map(s => ({
    value: s._id,
    label: `${s.name}${s.phone ? ` - ${s.phone}` : ''}`
  }))

  const itemOptions = items.map(item => ({
    value: item._id,
    label: `${item.name} (Stock: ${item.currentStock || 0} ${item.unit})`
  }))

  const accountOptions = accounts.map(acc => ({
    value: acc._id,
    label: `${acc.name} (${acc.code}) - ${formatCurrency(acc.currentBalance || 0)}`
  }))

  const paymentMethodOptions = [
    { value: 'cash', label: 'Cash' },
    { value: 'bank', label: 'Bank' },
    { value: 'cheque', label: 'Cheque' },
    { value: 'online', label: 'Online' },
  ]

  const typeOptions = [
    { value: '', label: 'All Types' },
    { value: 'sale', label: 'Sale' },
    { value: 'purchase', label: 'Purchase' },
  ]

  const paymentStatusOptions = [
    { value: '', label: 'All Status' },
    { value: 'pending', label: 'Pending' },
    { value: 'partial', label: 'Partial' },
    { value: 'paid', label: 'Paid' },
    { value: 'overdue', label: 'Overdue' },
  ]

  const paymentMethodFilterOptions = [
    { value: '', label: 'All Methods' },
    { value: 'cash', label: 'Cash' },
    { value: 'bank', label: 'Bank' },
    { value: 'cheque', label: 'Cheque' },
    { value: 'online', label: 'Online' },
  ]

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900">Transactions</h1>
          <p className="text-gray-600 mt-1.5">Manage sales and purchases</p>
        </div>
        <div className="flex gap-2">
          <Button
            variant="success"
            onClick={() => handleCreate('sale')}
            size="lg"
          >
            <ShoppingCart className="h-5 w-5 mr-2" />
            New Sale
          </Button>
          <Button
            variant="primary"
            onClick={() => handleCreate('purchase')}
            size="lg"
          >
            <Package className="h-5 w-5 mr-2" />
            New Purchase
          </Button>
        </div>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>All Transactions</CardTitle>
          <CardDescription>
            View and manage all your sales and purchases
          </CardDescription>
        </CardHeader>
        <CardContent>
          {/* Search and Filter Bar */}
          <div className="mb-6 space-y-4">
            <div className="flex gap-4">
              <div className="relative flex-1">
                <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
                <input
                  type="text"
                  placeholder="Search by transaction number, customer, supplier..."
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value)
                    setPagination(prev => ({ ...prev, page: 1 }))
                  }}
                  className="w-full pl-10 pr-4 py-2.5 border border-gray-300 bg-white rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-primary-500 focus:ring-offset-2 transition-all"
                />
              </div>
              <Button
                variant={showFilters ? 'primary' : 'outline'}
                onClick={() => setShowFilters(!showFilters)}
              >
                <Filter className="h-5 w-5 mr-2" />
                Filters
                {activeFiltersCount > 0 && (
                  <span className="ml-2 bg-white text-primary-600 rounded-full px-2 py-0.5 text-xs font-bold">
                    {activeFiltersCount}
                  </span>
                )}
              </Button>
              {activeFiltersCount > 0 && (
                <Button
                  variant="secondary"
                  onClick={clearFilters}
                >
                  <X className="h-5 w-5 mr-2" />
                  Clear
                </Button>
              )}
            </div>

            {/* Filter Panel */}
            {showFilters && (
              <div className="border-t pt-4">
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Transaction Type
                    </label>
                    <select
                      value={filters.type}
                      onChange={(e) => handleFilterChange('type', e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                    >
                      {typeOptions.map(option => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Payment Status
                    </label>
                    <select
                      value={filters.paymentStatus}
                      onChange={(e) => handleFilterChange('paymentStatus', e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                    >
                      {paymentStatusOptions.map(option => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Payment Method
                    </label>
                    <select
                      value={filters.paymentMethod}
                      onChange={(e) => handleFilterChange('paymentMethod', e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                    >
                      {paymentMethodFilterOptions.map(option => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Start Date
                    </label>
                    <input
                      type="date"
                      value={filters.startDate}
                      onChange={(e) => handleFilterChange('startDate', e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                    />
                  </div>

                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      End Date
                    </label>
                    <input
                      type="date"
                      value={filters.endDate}
                      onChange={(e) => handleFilterChange('endDate', e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                    />
                  </div>
                </div>
              </div>
            )}
          </div>

          <DataTable
            columns={columns}
            data={transactions}
            loading={loading}
            onView={handleView}
            onDelete={handleDelete}
          />
        </CardContent>
      </Card>

      {/* Simplified Sale/Purchase Modal */}
      <Dialog open={isModalOpen} onOpenChange={setIsModalOpen}>
        <DialogContent className="sm:max-w-[700px] max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>
              {transactionType === 'sale' ? 'Create New Sale' : 'Create New Purchase'}
            </DialogTitle>
            <DialogDescription>
              {transactionType === 'sale' 
                ? 'Record a sale to a customer'
                : 'Record a purchase from a supplier'}
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(onSubmit)} className="space-y-4">
            {/* Basic Info */}
            <div className="grid grid-cols-2 gap-4">
              <FormInput
                label="Date"
                name="date"
                type="date"
                register={register}
                required
                error={errors.date?.message}
              />
              <FormSelect
                label={transactionType === 'sale' ? 'Customer *' : 'Supplier *'}
                name={transactionType === 'sale' ? 'customer' : 'supplier'}
                register={register}
                required
                options={transactionType === 'sale' ? customerOptions : supplierOptions}
                error={errors[transactionType === 'sale' ? 'customer' : 'supplier']?.message}
                placeholder={`Select ${transactionType === 'sale' ? 'customer' : 'supplier'}`}
              />
            </div>

            {/* Simple Item Selection */}
            <div className="border-t pt-4">
              <h3 className="text-sm font-semibold text-gray-700 mb-3">Add Items</h3>
              <div className="flex gap-2 mb-4">
                <div className="flex-1">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Select Item
                  </label>
                  <select
                    value={selectedItem?._id || ''}
                    onChange={(e) => {
                      const item = items.find(i => i._id === e.target.value)
                      setSelectedItem(item || null)
                    }}
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                  >
                    <option value="">Choose an item...</option>
                    {itemOptions.map(option => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="w-32">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Quantity
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    min="0.01"
                    value={itemQuantity}
                    onChange={(e) => setItemQuantity(e.target.value)}
                    placeholder="Qty"
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                  />
                </div>
                <div className="flex items-end">
                  <Button
                    type="button"
                    onClick={addItemToTransaction}
                    disabled={!selectedItem}
                    variant="outline"
                  >
                    <Plus className="h-4 w-4 mr-1" />
                    Add
                  </Button>
                </div>
              </div>

              {/* Items List */}
              {transactionItems.length > 0 ? (
                <div className="space-y-2 border rounded-lg p-3 bg-gray-50">
                  {transactionItems.map((ti, index) => (
                    <div key={index} className="flex items-center justify-between p-2 bg-white rounded border">
                      <div className="flex-1">
                        <p className="font-medium text-sm">{ti.item.name}</p>
                        <p className="text-xs text-gray-500">
                          Stock: {ti.item.currentStock || 0} {ti.item.unit} | 
                          Rate: {formatCurrency(ti.rate)}
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <input
                          type="number"
                          min="0.01"
                          step="0.01"
                          value={ti.quantity}
                          onChange={(e) => updateItemQuantity(index, e.target.value)}
                          className="w-20 px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-primary-500"
                        />
                        <span className="text-sm font-semibold w-24 text-right">
                          {formatCurrency(ti.total)}
                        </span>
                        <Button
                          type="button"
                          variant="ghost"
                          size="sm"
                          onClick={() => removeItem(index)}
                        >
                          <Trash2 className="h-4 w-4 text-red-600" />
                        </Button>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-6 text-gray-400 text-sm border border-dashed rounded-lg">
                  No items added yet
                </div>
              )}
            </div>

            {/* Totals */}
            <div className="border-t pt-4 space-y-3">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <p className="text-sm font-medium text-gray-600 mb-1">Subtotal</p>
                  <p className="text-lg font-bold text-gray-900">{formatCurrency(totals.subtotal)}</p>
                </div>
                <FormInput
                  label="Discount"
                  name="discount"
                  type="number"
                  step="0.01"
                  register={register}
                  error={errors.discount?.message}
                  placeholder="0.00"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <FormInput
                  label="Tax"
                  name="tax"
                  type="number"
                  step="0.01"
                  register={register}
                  error={errors.tax?.message}
                  placeholder="0.00"
                />
                <div>
                  <p className="text-sm font-medium text-gray-600 mb-1">Total</p>
                  <p className="text-lg font-bold text-primary-600">{formatCurrency(totals.total)}</p>
                </div>
              </div>
            </div>

            {/* Payment Section */}
            <div className="border-t pt-4 space-y-3">
              <h3 className="text-sm font-semibold text-gray-700">Payment Details</h3>
              <div className="grid grid-cols-2 gap-4">
                <FormSelect
                  label={transactionType === 'sale' ? 'Receive Money To Account' : 'Pay Money From Account'}
                  name="account"
                  register={register}
                  options={accountOptions}
                  error={errors.account?.message}
                  placeholder={`Select account${transactionType === 'sale' ? ' to receive money' : ' to pay from'}`}
                />
                <FormSelect
                  label="Payment Method"
                  name="paymentMethod"
                  register={register}
                  required
                  options={paymentMethodOptions}
                  error={errors.paymentMethod?.message}
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <FormInput
                  label="Paid Amount"
                  name="paidAmount"
                  type="number"
                  step="0.01"
                  register={register}
                  error={errors.paidAmount?.message}
                  placeholder="0.00"
                />
                <div>
                  <p className="text-sm font-medium text-gray-600 mb-1">Remaining</p>
                  <p className={`text-lg font-bold ${totals.remaining > 0 ? 'text-red-600' : 'text-green-600'}`}>
                    {formatCurrency(totals.remaining)}
                  </p>
                </div>
              </div>
            </div>

            <FormInput
              label="Notes (Optional)"
              name="notes"
              register={register}
              error={errors.notes?.message}
              placeholder="Additional notes..."
            />

            <DialogFooter>
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  setIsModalOpen(false)
                  reset()
                  setTransactionItems([])
                  setSelectedItem(null)
                }}
              >
                Cancel
              </Button>
              <Button type="submit" variant={transactionType === 'sale' ? 'success' : 'primary'}>
                Create {transactionType === 'sale' ? 'Sale' : 'Purchase'}
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* View Transaction Modal */}
      <Dialog open={isViewModalOpen} onOpenChange={setIsViewModalOpen}>
        <DialogContent className="sm:max-w-[700px] max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>
              Transaction Details - {viewingTransaction?.transactionNumber}
            </DialogTitle>
            <DialogDescription>
              View complete transaction information
            </DialogDescription>
          </DialogHeader>
          {viewingTransaction && (
            <div className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <p className="text-sm font-medium text-gray-500">Type</p>
                  <p className="text-base font-semibold capitalize">{viewingTransaction.type}</p>
                </div>
                <div>
                  <p className="text-sm font-medium text-gray-500">Date</p>
                  <p className="text-base font-semibold">{formatDate(viewingTransaction.date)}</p>
                </div>
                {viewingTransaction.customer && (
                  <div>
                    <p className="text-sm font-medium text-gray-500">Customer</p>
                    <p className="text-base font-semibold">{viewingTransaction.customer.name}</p>
                  </div>
                )}
                {viewingTransaction.supplier && (
                  <div>
                    <p className="text-sm font-medium text-gray-500">Supplier</p>
                    <p className="text-base font-semibold">{viewingTransaction.supplier.name}</p>
                  </div>
                )}
              </div>

              <div className="border-t pt-4">
                <h4 className="font-semibold mb-3">Items</h4>
                <div className="space-y-2">
                  {viewingTransaction.items?.map((item, index) => (
                    <div key={index} className="flex justify-between items-center p-2 bg-gray-50 rounded">
                      <div>
                        <p className="font-medium">{item.item?.name || 'N/A'}</p>
                        <p className="text-sm text-gray-500">
                          {item.quantity} {item.item?.unit || ''} × {formatCurrency(item.rate)}
                        </p>
                      </div>
                      <p className="font-semibold">{formatCurrency(item.total)}</p>
                    </div>
                  ))}
                </div>
              </div>

              <div className="border-t pt-4 space-y-2">
                <div className="flex justify-between">
                  <span className="text-gray-600">Subtotal:</span>
                  <span className="font-semibold">{formatCurrency(viewingTransaction.subtotal)}</span>
                </div>
                {viewingTransaction.discount > 0 && (
                  <div className="flex justify-between">
                    <span className="text-gray-600">Discount:</span>
                    <span className="font-semibold text-red-600">-{formatCurrency(viewingTransaction.discount)}</span>
                  </div>
                )}
                {viewingTransaction.tax > 0 && (
                  <div className="flex justify-between">
                    <span className="text-gray-600">Tax:</span>
                    <span className="font-semibold">{formatCurrency(viewingTransaction.tax)}</span>
                  </div>
                )}
                <div className="flex justify-between border-t pt-2">
                  <span className="text-lg font-bold text-gray-900">Total:</span>
                  <span className="text-lg font-bold text-primary-600">{formatCurrency(viewingTransaction.total)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-600">Paid:</span>
                  <span className="font-semibold text-green-600">{formatCurrency(viewingTransaction.paidAmount)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-600">Remaining:</span>
                  <span className={`font-semibold ${viewingTransaction.remainingAmount > 0 ? 'text-red-600' : 'text-green-600'}`}>
                    {formatCurrency(viewingTransaction.remainingAmount)}
                  </span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-600">Payment Status:</span>
                  <span className={`px-2 py-1 rounded-full text-xs font-semibold ${
                    viewingTransaction.paymentStatus === 'paid' ? 'bg-green-100 text-green-800' :
                    viewingTransaction.paymentStatus === 'partial' ? 'bg-orange-100 text-orange-800' :
                    'bg-yellow-100 text-yellow-800'
                  }`}>
                    {viewingTransaction.paymentStatus?.toUpperCase()}
                  </span>
                </div>
              </div>

              {viewingTransaction.notes && (
                <div className="border-t pt-4">
                  <p className="text-sm font-medium text-gray-500 mb-1">Notes</p>
                  <p className="text-sm text-gray-700">{viewingTransaction.notes}</p>
                </div>
              )}
            </div>
          )}
          <DialogFooter>
            <Button variant="outline" onClick={() => setIsViewModalOpen(false)}>
              Close
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  )
}

export default Transactions
