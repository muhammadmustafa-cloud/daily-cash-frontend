import { useState } from 'react'
import { Plus, Search, FileText, ArrowRight } from 'lucide-react'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Table from '../../components/ui/Table'
import { formatCurrency, formatDate } from '../../utils/formatters'
import { useNavigate } from 'react-router-dom'

const Vouchers = () => {
  const navigate = useNavigate()
  const [vouchers, setVouchers] = useState([
    {
      _id: '1',
      voucherNumber: 'BP-20240115-001',
      voucherType: 'bank_payment',
      date: new Date('2024-01-15'),
      description: 'Payment to Supplier ABC',
      totalAmount: 50000,
      status: 'posted',
      account: 'Bank - HBL Main',
    },
    {
      _id: '2',
      voucherNumber: 'BR-20240116-001',
      voucherType: 'bank_receipt',
      date: new Date('2024-01-16'),
      description: 'Payment received from Customer XYZ',
      totalAmount: 75000,
      status: 'posted',
      account: 'Bank - HBL Main',
    },
    {
      _id: '3',
      voucherNumber: 'CP-20240117-001',
      voucherType: 'cash_payment',
      date: new Date('2024-01-17'),
      description: 'Mazdoor payment',
      totalAmount: 15000,
      status: 'posted',
      account: 'Cash Account',
    },
    {
      _id: '4',
      voucherNumber: 'CR-20240118-001',
      voucherType: 'cash_receipt',
      date: new Date('2024-01-18'),
      description: 'Cash sale',
      totalAmount: 25000,
      status: 'posted',
      account: 'Cash Account',
    },
    {
      _id: '5',
      voucherNumber: 'BT-20240119-001',
      voucherType: 'bank_transfer',
      date: new Date('2024-01-19'),
      description: 'Transfer from HBL to UBL',
      totalAmount: 100000,
      status: 'posted',
      account: 'Bank Transfer',
    },
    {
      _id: '6',
      voucherNumber: 'JE-20240120-001',
      voucherType: 'journal_entry',
      date: new Date('2024-01-20'),
      description: 'Adjustment entry',
      totalAmount: 5000,
      status: 'draft',
      account: 'Journal Entry',
    },
  ])

  const [searchTerm, setSearchTerm] = useState('')
  const [filterType, setFilterType] = useState('all')

  const filteredVouchers = vouchers.filter(voucher => {
    const matchesSearch = 
      voucher.voucherNumber.toLowerCase().includes(searchTerm.toLowerCase()) ||
      voucher.description.toLowerCase().includes(searchTerm.toLowerCase())
    const matchesType = filterType === 'all' || voucher.voucherType === filterType
    return matchesSearch && matchesType
  })

  const getVoucherTypeLabel = (type) => {
    const labels = {
      bank_payment: 'Bank Payment',
      bank_receipt: 'Bank Receipt',
      cash_payment: 'Cash Payment',
      cash_receipt: 'Cash Receipt',
      bank_transfer: 'Bank Transfer',
      journal_entry: 'Journal Entry',
    }
    return labels[type] || type
  }

  const getVoucherTypeColor = (type) => {
    const colors = {
      bank_payment: 'bg-red-100 text-red-800',
      bank_receipt: 'bg-green-100 text-green-800',
      cash_payment: 'bg-orange-100 text-orange-800',
      cash_receipt: 'bg-blue-100 text-blue-800',
      bank_transfer: 'bg-purple-100 text-purple-800',
      journal_entry: 'bg-gray-100 text-gray-800',
    }
    return colors[type] || 'bg-gray-100 text-gray-800'
  }

  const columns = [
    { key: 'voucherNumber', label: 'Voucher #' },
    {
      key: 'voucherType',
      label: 'Type',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs ${getVoucherTypeColor(value)}`}>
          {getVoucherTypeLabel(value)}
        </span>
      ),
    },
    { key: 'description', label: 'Description' },
    {
      key: 'totalAmount',
      label: 'Amount',
      render: (value) => formatCurrency(value || 0),
    },
    { key: 'account', label: 'Account' },
    {
      key: 'date',
      label: 'Date',
      render: (value) => formatDate(value),
    },
    {
      key: 'status',
      label: 'Status',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs ${
          value === 'posted' ? 'bg-green-100 text-green-800' :
          value === 'draft' ? 'bg-yellow-100 text-yellow-800' :
          'bg-red-100 text-red-800'
        }`}>
          {value?.toUpperCase() || 'DRAFT'}
        </span>
      ),
    },
  ]

  const voucherTypeOptions = [
    { value: 'all', label: 'All Types' },
    { value: 'bank_payment', label: 'Bank Payment' },
    { value: 'bank_receipt', label: 'Bank Receipt' },
    { value: 'cash_payment', label: 'Cash Payment' },
    { value: 'cash_receipt', label: 'Cash Receipt' },
    { value: 'bank_transfer', label: 'Bank Transfer' },
    { value: 'journal_entry', label: 'Journal Entry' },
  ]

  return (
    <div>
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Vouchers</h1>
          <p className="text-gray-600 mt-1">Manage all financial vouchers</p>
        </div>
        <div className="flex gap-2 flex-wrap">
          <Button
            variant="primary"
            size="sm"
            onClick={() => navigate('/vouchers/bank-payment')}
          >
            <Plus className="h-5 w-5 mr-2" />
            Bank Payment
          </Button>
          <Button
            variant="primary"
            size="sm"
            onClick={() => navigate('/vouchers/bank-receipt')}
          >
            <Plus className="h-5 w-5 mr-2" />
            Bank Receipt
          </Button>
          <Button
            variant="primary"
            size="sm"
            onClick={() => navigate('/vouchers/cash-payment')}
          >
            <Plus className="h-5 w-5 mr-2" />
            Cash Payment
          </Button>
          <Button
            variant="primary"
            size="sm"
            onClick={() => navigate('/vouchers/cash-receipt')}
          >
            <Plus className="h-5 w-5 mr-2" />
            Cash Receipt
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={() => navigate('/vouchers/bank-transfer')}
          >
            <Plus className="h-5 w-5 mr-2" />
            Bank Transfer
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={() => navigate('/vouchers/journal-entry')}
          >
            <Plus className="h-5 w-5 mr-2" />
            Journal Entry
          </Button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
        <Card>
          <div className="text-center">
            <p className="text-sm text-gray-600">Bank Payments</p>
            <p className="text-2xl font-bold text-gray-900">
              {vouchers.filter(v => v.voucherType === 'bank_payment').length}
            </p>
          </div>
        </Card>
        <Card>
          <div className="text-center">
            <p className="text-sm text-gray-600">Bank Receipts</p>
            <p className="text-2xl font-bold text-gray-900">
              {vouchers.filter(v => v.voucherType === 'bank_receipt').length}
            </p>
          </div>
        </Card>
        <Card>
          <div className="text-center">
            <p className="text-sm text-gray-600">Cash Payments</p>
            <p className="text-2xl font-bold text-gray-900">
              {vouchers.filter(v => v.voucherType === 'cash_payment').length}
            </p>
          </div>
        </Card>
        <Card>
          <div className="text-center">
            <p className="text-sm text-gray-600">Cash Receipts</p>
            <p className="text-2xl font-bold text-gray-900">
              {vouchers.filter(v => v.voucherType === 'cash_receipt').length}
            </p>
          </div>
        </Card>
      </div>

      <Card>
        <div className="mb-4 flex gap-4">
          <div className="relative flex-1">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search vouchers..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
            />
          </div>
          <select
            value={filterType}
            onChange={(e) => setFilterType(e.target.value)}
            className="px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
          >
            {voucherTypeOptions.map(option => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        <Table columns={columns} data={filteredVouchers} />
      </Card>
    </div>
  )
}

export default Vouchers

