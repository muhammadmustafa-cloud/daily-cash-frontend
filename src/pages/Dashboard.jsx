import { useEffect, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { 
  Users, 
  ShoppingBag, 
  Package, 
  DollarSign, 
  TrendingUp,
  TrendingDown,
  AlertCircle,
  ArrowRight
} from 'lucide-react'
import Card from '../components/ui/Card'
import api from '../config/api'
import toast from 'react-hot-toast'
import { formatCurrency } from '../utils/formatters'

const Dashboard = () => {
  const [stats, setStats] = useState({
    totalCustomers: 0,
    totalSuppliers: 0,
    totalItems: 0,
    totalExpenses: 0,
    totalRevenue: 0,
    lowStockItems: 0,
  })
  const [loading, setLoading] = useState(true)
  const navigate = useNavigate()

  useEffect(() => {
    fetchDashboardStats()
  }, [])

  const fetchDashboardStats = async () => {
    try {
      setLoading(true)
      const [customersRes, suppliersRes, itemsRes, expensesRes, transactionsRes] = await Promise.all([
        api.get('/customers?limit=1'),
        api.get('/suppliers?limit=1'),
        api.get('/items?limit=1'),
        api.get('/expenses?limit=1'),
        api.get('/transactions?limit=1'),
      ])

      const itemsData = await api.get('/items?limit=1000')
      const lowStockCount = itemsData.data.data.filter(item => 
        item.minStockLevel > 0 && item.currentStock <= item.minStockLevel
      ).length

      const expensesData = await api.get('/expenses?limit=1000')
      const totalExpenses = expensesData.data.data.reduce((sum, exp) => sum + (exp.amount || 0), 0)

      const transactionsData = await api.get('/transactions?limit=1000')
      const totalRevenue = transactionsData.data.data
        .filter(t => t.type === 'sale')
        .reduce((sum, t) => sum + (t.totalAmount || 0), 0)

      setStats({
        totalCustomers: customersRes.data.pagination?.total || 0,
        totalSuppliers: suppliersRes.data.pagination?.total || 0,
        totalItems: itemsRes.data.pagination?.total || 0,
        totalExpenses,
        totalRevenue,
        lowStockItems: lowStockCount,
      })
    } catch (error) {
      console.error('Error fetching dashboard stats:', error)
      toast.error('Failed to load dashboard data')
    } finally {
      setLoading(false)
    }
  }

  const statCards = [
    {
      title: 'Total Customers',
      value: stats.totalCustomers,
      icon: Users,
      gradient: 'from-blue-500 to-blue-600',
      bgGradient: 'from-blue-50 to-blue-100',
      onClick: () => navigate('/customers'),
    },
    {
      title: 'Total Suppliers',
      value: stats.totalSuppliers,
      icon: ShoppingBag,
      gradient: 'from-green-500 to-green-600',
      bgGradient: 'from-green-50 to-green-100',
      onClick: () => navigate('/suppliers'),
    },
    {
      title: 'Total Items',
      value: stats.totalItems,
      icon: Package,
      gradient: 'from-purple-500 to-purple-600',
      bgGradient: 'from-purple-50 to-purple-100',
      onClick: () => navigate('/items'),
    },
    {
      title: 'Total Revenue',
      value: formatCurrency(stats.totalRevenue),
      icon: TrendingUp,
      gradient: 'from-emerald-500 to-emerald-600',
      bgGradient: 'from-emerald-50 to-emerald-100',
      onClick: () => navigate('/transactions'),
    },
    {
      title: 'Total Expenses',
      value: formatCurrency(stats.totalExpenses),
      icon: TrendingDown,
      gradient: 'from-red-500 to-red-600',
      bgGradient: 'from-red-50 to-red-100',
      onClick: () => navigate('/expenses'),
    },
    {
      title: 'Low Stock Items',
      value: stats.lowStockItems,
      icon: AlertCircle,
      gradient: 'from-orange-500 to-orange-600',
      bgGradient: 'from-orange-50 to-orange-100',
      onClick: () => navigate('/inventory'),
    },
  ]

  const quickActions = [
    {
      title: 'Add Customer',
      icon: Users,
      path: '/customers',
      color: 'text-blue-600 bg-blue-50 hover:bg-blue-100',
    },
    {
      title: 'Add Supplier',
      icon: ShoppingBag,
      path: '/suppliers',
      color: 'text-green-600 bg-green-50 hover:bg-green-100',
    },
    {
      title: 'Stock Entry',
      icon: Package,
      path: '/inventory',
      color: 'text-purple-600 bg-purple-50 hover:bg-purple-100',
    },
    {
      title: 'Add Expense',
      icon: DollarSign,
      path: '/expenses',
      color: 'text-red-600 bg-red-50 hover:bg-red-100',
    },
  ]

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-primary-600"></div>
      </div>
    )
  }

  return (
    <div className="space-y-6">
      <div className="mb-8">
        <h1 className="text-4xl font-bold text-gray-900 mb-2">Dashboard</h1>
        <p className="text-lg text-gray-600">Welcome to Daily Cash Memo</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {statCards.map((stat, index) => {
          const Icon = stat.icon
          return (
            <Card
              key={index}
              className="cursor-pointer hover:shadow-xl transition-all duration-300 hover:scale-[1.02] group overflow-hidden relative"
              onClick={stat.onClick}
            >
              <div className={`absolute inset-0 bg-gradient-to-br ${stat.bgGradient} opacity-0 group-hover:opacity-100 transition-opacity duration-300`}></div>
              <div className="relative z-10 flex items-center justify-between p-6">
                <div>
                  <p className="text-sm font-semibold text-gray-600 mb-2">{stat.title}</p>
                  <p className="text-3xl font-bold text-gray-900">{stat.value}</p>
                </div>
                <div className={`bg-gradient-to-br ${stat.gradient} p-4 rounded-xl shadow-lg group-hover:scale-110 transition-transform duration-300`}>
                  <Icon className="h-7 w-7 text-white" />
                </div>
              </div>
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-gradient-to-r from-transparent via-primary-500 to-transparent opacity-0 group-hover:opacity-100 transition-opacity"></div>
            </Card>
          )
        })}
      </div>

      <Card className="bg-gradient-to-br from-gray-50 to-white">
        <div className="p-6">
          <h2 className="text-2xl font-bold text-gray-900 mb-6">Quick Actions</h2>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            {quickActions.map((action, index) => {
              const Icon = action.icon
              return (
                <button
                  key={index}
                  onClick={() => navigate(action.path)}
                  className={`p-6 rounded-xl border-2 border-gray-200 ${action.color} transition-all duration-200 hover:scale-105 hover:shadow-lg group text-left`}
                >
                  <Icon className="h-8 w-8 mb-3 group-hover:scale-110 transition-transform" />
                  <p className="font-semibold text-sm">{action.title}</p>
                  <ArrowRight className="h-4 w-4 mt-2 opacity-0 group-hover:opacity-100 transition-opacity" />
                </button>
              )
            })}
          </div>
        </div>
      </Card>
    </div>
  )
}

export default Dashboard
