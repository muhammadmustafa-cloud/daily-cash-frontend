import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Search, Building2, Wallet } from 'lucide-react'
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
import { formatCurrency } from '../../utils/formatters'

const Accounts = () => {
  const [accounts, setAccounts] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingAccount, setEditingAccount] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [stats, setStats] = useState({
    totalCash: 0,
    totalBank: 0,
    totalAccounts: 0
  })

  const { register, handleSubmit, formState: { errors }, reset, setValue, watch } = useForm()
  const isBankAccount = watch('isBankAccount')
  const accountType = watch('type')

  useEffect(() => {
    fetchAccounts()
  }, [searchTerm])

  const fetchAccounts = async () => {
    try {
      setLoading(true)
      const response = await api.get('/accounts', {
        params: {
          page: 1,
          limit: 1000,
          search: searchTerm,
          isActive: true,
        },
      })
      setAccounts(response.data.data || [])
      
      // Update stats from pagination data
      if (response.data.pagination) {
        setStats({
          totalCash: response.data.pagination.totalCash || 0,
          totalBank: response.data.pagination.totalBank || 0,
          totalAccounts: response.data.pagination.total || 0
        })
      }
    } catch (error) {
      toast.error('Failed to fetch accounts')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    setEditingAccount(null)
    reset({
      type: 'asset',
      isCashAccount: false,
      isBankAccount: false,
      openingBalance: 0,
    })
    setIsModalOpen(true)
  }

  const handleEdit = (account) => {
    setEditingAccount(account)
    setValue('code', account.code)
    setValue('name', account.name)
    setValue('type', account.type)
    setValue('openingBalance', account.openingBalance || 0)
    setValue('isCashAccount', account.isCashAccount || false)
    setValue('isBankAccount', account.isBankAccount || false)
    setValue('bankDetails.bankName', account.bankDetails?.bankName || '')
    setValue('bankDetails.accountNumber', account.bankDetails?.accountNumber || '')
    setValue('bankDetails.branch', account.bankDetails?.branch || '')
    setValue('notes', account.notes || '')
    setIsModalOpen(true)
  }

  const handleDelete = async (account) => {
    if (!window.confirm(`Are you sure you want to delete ${account.name}?`)) {
      return
    }

    try {
      await api.delete(`/accounts/${account._id}`)
      toast.success('Account deleted successfully')
      fetchAccounts()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete account')
    }
  }

  const onSubmit = async (data) => {
    try {
      const accountData = {
        ...data,
        bankDetails: data.isBankAccount ? {
          bankName: data['bankDetails.bankName'],
          accountNumber: data['bankDetails.accountNumber'],
          branch: data['bankDetails.branch'],
        } : undefined,
      }
      
      // Clean up the data
      delete accountData['bankDetails.bankName']
      delete accountData['bankDetails.accountNumber']
      delete accountData['bankDetails.branch']

      if (editingAccount) {
        await api.put(`/accounts/${editingAccount._id}`, accountData)
        toast.success('Account updated successfully')
      } else {
        await api.post('/accounts', accountData)
        toast.success('Account created successfully')
      }
      setIsModalOpen(false)
      reset()
      fetchAccounts()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save account')
    }
  }

  const columns = [
    { 
      key: 'code', 
      label: 'Code',
    },
    { 
      key: 'name', 
      label: 'Account Name',
    },
    {
      key: 'type',
      label: 'Type',
      render: (value) => (
        <span className="px-2 py-1 rounded-full text-xs font-semibold bg-blue-100 text-blue-800 capitalize">
          {value}
        </span>
      ),
    },
    {
      key: 'currentBalance',
      label: 'Balance',
      render: (value) => formatCurrency(value || 0),
    },
    {
      key: 'isBankAccount',
      label: 'Bank',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs ${value ? 'bg-green-100 text-green-800' : 'bg-gray-100 text-gray-600'}`}>
          {value ? 'Yes' : 'No'}
        </span>
      ),
    },
    {
      key: 'isCashAccount',
      label: 'Cash',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs ${value ? 'bg-green-100 text-green-800' : 'bg-gray-100 text-gray-600'}`}>
          {value ? 'Yes' : 'No'}
        </span>
      ),
    },
    {
      key: 'isActive',
      label: 'Status',
      render: (value) => (
        <span className={`px-2 py-1 rounded-full text-xs font-semibold ${
          value ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'
        }`}>
          {value ? 'Active' : 'Inactive'}
        </span>
      ),
    },
  ]

  const accountTypeOptions = [
    { value: 'asset', label: 'Asset' },
    { value: 'liability', label: 'Liability' },
    { value: 'income', label: 'Income' },
    { value: 'expense', label: 'Expense' },
    { value: 'equity', label: 'Equity' },
  ]

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900">Accounts</h1>
          <p className="text-gray-600 mt-1.5">Manage Chart of Accounts</p>
        </div>
        <Button onClick={handleCreate} size="lg">
          <Plus className="h-5 w-5 mr-2" />
          Add Account
        </Button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm font-medium text-gray-600">Total Cash</p>
                <p className="text-2xl font-bold text-gray-900 mt-1">
                  {formatCurrency(stats.totalCash)}
                </p>
              </div>
              <Wallet className="h-8 w-8 text-green-500" />
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm font-medium text-gray-600">Total Bank Balance</p>
                <p className="text-2xl font-bold text-gray-900 mt-1">
                  {formatCurrency(stats.totalBank)}
                </p>
              </div>
              <Building2 className="h-8 w-8 text-blue-500" />
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm font-medium text-gray-600">Total Accounts</p>
                <p className="text-2xl font-bold text-gray-900 mt-1">{stats.totalAccounts}</p>
              </div>
              <Building2 className="h-8 w-8 text-purple-500" />
            </div>
          </CardContent>
        </Card>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>All Accounts</CardTitle>
          <CardDescription>
            View and manage all your accounts
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="mb-6">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
              <input
                type="text"
                placeholder="Search accounts..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2.5 border border-gray-300 bg-white rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-primary-500 focus:ring-offset-2 transition-all"
              />
            </div>
          </div>
          <DataTable
            columns={columns}
            data={accounts}
            loading={loading}
            onEdit={handleEdit}
            onDelete={handleDelete}
          />
        </CardContent>
      </Card>

      <Dialog open={isModalOpen} onOpenChange={setIsModalOpen}>
        <DialogContent className="sm:max-w-[600px] max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>
              {editingAccount ? 'Edit Account' : 'Add New Account'}
            </DialogTitle>
            <DialogDescription>
              {editingAccount 
                ? 'Update account information below.' 
                : 'Enter account details to add it to your chart of accounts.'}
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(onSubmit)} className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <FormInput
                label="Account Code"
                name="code"
                register={register}
                required={!editingAccount}
                disabled={!!editingAccount}
                error={errors.code?.message}
                placeholder="e.g., 1001"
              />
              <FormInput
                label="Account Name"
                name="name"
                register={register}
                required
                error={errors.name?.message}
                placeholder="Account name"
              />
            </div>
            
            <FormSelect
              label="Account Type"
              name="type"
              register={register}
              required
              options={accountTypeOptions}
              error={errors.type?.message}
            />
            
            <FormInput
              label="Opening Balance"
              name="openingBalance"
              type="number"
              step="0.01"
              register={register}
              error={errors.openingBalance?.message}
              placeholder="0.00"
            />
            
            <div className="flex gap-6">
              <div className="flex items-center space-x-2">
                <input
                  type="checkbox"
                  id="isCashAccount"
                  {...register('isCashAccount')}
                  className="h-4 w-4 rounded border-gray-300 text-primary-600 focus:ring-primary-500"
                />
                <label htmlFor="isCashAccount" className="text-sm font-medium leading-none text-gray-700">
                  Cash Account
                </label>
              </div>
              <div className="flex items-center space-x-2">
                <input
                  type="checkbox"
                  id="isBankAccount"
                  {...register('isBankAccount')}
                  className="h-4 w-4 rounded border-gray-300 text-primary-600 focus:ring-primary-500"
                />
                <label htmlFor="isBankAccount" className="text-sm font-medium leading-none text-gray-700">
                  Bank Account
                </label>
              </div>
            </div>
            
            {isBankAccount && (
              <div className="space-y-4 border-t pt-4">
                <FormInput
                  label="Bank Name"
                  name="bankDetails.bankName"
                  register={register}
                  error={errors['bankDetails.bankName']?.message}
                  placeholder="Bank name"
                />
                <FormInput
                  label="Account Number"
                  name="bankDetails.accountNumber"
                  register={register}
                  error={errors['bankDetails.accountNumber']?.message}
                  placeholder="Account number"
                />
                <FormInput
                  label="Branch"
                  name="bankDetails.branch"
                  register={register}
                  error={errors['bankDetails.branch']?.message}
                  placeholder="Branch name"
                />
              </div>
            )}
            
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
                }}
              >
                Cancel
              </Button>
              <Button type="submit">
                {editingAccount ? 'Update' : 'Create'} Account
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}

export default Accounts
