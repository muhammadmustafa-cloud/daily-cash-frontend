import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Search } from 'lucide-react'
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
import { formatDate, formatPhone } from '../../utils/formatters'

const Customers = () => {
  const [customers, setCustomers] = useState([])
  const [loading, setLoading] = useState(false)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingCustomer, setEditingCustomer] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')

  const { register, handleSubmit, formState: { errors }, reset, setValue } = useForm()

  useEffect(() => {
    fetchCustomers()
  }, [searchTerm])

  const fetchCustomers = async () => {
    try {
      setLoading(true)
      const response = await api.get('/customers', {
        params: {
          page: 1,
          limit: 1000,
          search: searchTerm,
        },
      })
      setCustomers(response.data.data || [])
    } catch (error) {
      toast.error('Failed to fetch customers')
      console.error(error)
    } finally {
      setLoading(false)
    }
  }

  const handleCreate = () => {
    setEditingCustomer(null)
    reset()
    setIsModalOpen(true)
  }

  const handleEdit = (customer) => {
    setEditingCustomer(customer)
    setValue('name', customer.name)
    setValue('address', customer.address || '')
    setValue('phone', customer.phone || '')
    setValue('isActive', customer.isActive ?? true)
    setIsModalOpen(true)
  }

  const handleDelete = async (customer) => {
    if (!window.confirm(`Are you sure you want to delete ${customer.name}?`)) {
      return
    }

    try {
      await api.delete(`/customers/${customer._id}`)
      toast.success('Customer deleted successfully')
      fetchCustomers()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete customer')
    }
  }

  const onSubmit = async (data) => {
    try {
      if (editingCustomer) {
        await api.put(`/customers/${editingCustomer._id}`, data)
        toast.success('Customer updated successfully')
      } else {
        await api.post('/customers', data)
        toast.success('Customer created successfully')
      }
      setIsModalOpen(false)
      reset()
      fetchCustomers()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save customer')
    }
  }

  const columns = [
    { 
      key: 'name', 
      label: 'Name',
    },
    { 
      key: 'address', 
      label: 'Address',
    },
    { 
      key: 'phone', 
      label: 'Contact No', 
      render: (value) => value ? formatPhone(value) : '-' 
    },
    {
      key: 'isActive',
      label: 'Status',
      render: (value) => (
        <span className={`px-3 py-1 rounded-full text-xs font-semibold ${
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

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900">Customers</h1>
          <p className="text-gray-600 mt-1.5">Manage your customer database</p>
        </div>
        <Button onClick={handleCreate} size="lg">
          <Plus className="h-5 w-5 mr-2" />
          Add Customer
        </Button>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>All Customers</CardTitle>
          <CardDescription>
            View and manage all your customers
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="mb-6">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
              <input
                type="text"
                placeholder="Search customers..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2.5 border border-gray-300 bg-white rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-primary-500 focus:ring-offset-2 transition-all"
              />
            </div>
          </div>
          <DataTable
            columns={columns}
            data={customers}
            loading={loading}
            onEdit={handleEdit}
            onDelete={handleDelete}
          />
        </CardContent>
      </Card>

      <Dialog open={isModalOpen} onOpenChange={setIsModalOpen}>
        <DialogContent className="sm:max-w-[500px]">
          <DialogHeader>
            <DialogTitle>
              {editingCustomer ? 'Edit Customer' : 'Add New Customer'}
            </DialogTitle>
            <DialogDescription>
              {editingCustomer 
                ? 'Update customer information below.' 
                : 'Enter customer details to add them to your database.'}
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(onSubmit)} className="space-y-4">
            <FormInput
              label="Name"
              name="name"
              register={register}
              required
              error={errors.name?.message}
              placeholder="Customer name"
            />
            <FormInput
              label="Address"
              name="address"
              register={register}
              required
              error={errors.address?.message}
              placeholder="Customer address"
            />
            <FormInput
              label="Contact No (Optional)"
              name="phone"
              register={register}
              error={errors.phone?.message}
              placeholder="03XX-XXXXXXX"
            />
            <div className="flex items-center space-x-2">
              <input
                type="checkbox"
                id="isActive"
                {...register('isActive')}
                defaultChecked={true}
                className="h-4 w-4 rounded border-gray-300 text-primary-600 focus:ring-primary-500"
              />
              <label htmlFor="isActive" className="text-sm font-medium leading-none text-gray-700">
                Active
              </label>
            </div>
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
                {editingCustomer ? 'Update' : 'Create'} Customer
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}

export default Customers
