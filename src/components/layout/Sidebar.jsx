import { NavLink } from 'react-router-dom'
import {
  LayoutDashboard,
  Users,
  ShoppingBag,
  Package,
  TrendingUp,
  DollarSign,
  UserCheck,
  FileText,
  Settings,
  LogOut,
  Building2,
  CreditCard,
  Receipt,
  UsersIcon,
  Clock,
  Package2
} from 'lucide-react'
import { useAuth } from '../../contexts/AuthContext'
import { cn } from '../../lib/utils'
import Button from '../ui/Button'

const Sidebar = () => {
  const { logout, user } = useAuth()

  const menuGroups = [
    {
      title: 'Main',
      items: [
        { path: '/', label: 'Dashboard', icon: LayoutDashboard },
      ]
    },
    {
      title: 'People',
      items: [
        { path: '/customers', label: 'Customers', icon: Users },
        { path: '/suppliers', label: 'Suppliers', icon: ShoppingBag },
        { path: '/mazdoors', label: 'Mazdoors', icon: UserCheck },
      ]
    },

    {
      title: 'Inventory',
      items: [
        { path: '/items', label: 'Items', icon: Package },
        { path: '/inventory', label: 'Stock', icon: TrendingUp },
      ]
    },
    {
      title: 'Financial',
      items: [
        { path: '/transactions', label: 'Transactions', icon: FileText },
        { path: '/expenses', label: 'Expenses', icon: DollarSign },
        { path: '/payments', label: 'Payments', icon: CreditCard },
        { path: '/accounts', label: 'Accounts', icon: Building2 },
        { path: '/labour-expenses', label: 'Labour Expenses', icon: Clock },
        { path: '/labour', label: 'Labour', icon: UsersIcon },
        { path: '/daily-cash-memo', label: 'Daily Cash Memo', icon: Receipt },
      ]
    }
  ]

  // Add Users menu only for admin
  if (user?.role === 'admin') {
    menuGroups.push({
      title: 'Settings',
      items: [
        { path: '/users', label: 'Users', icon: Settings },
      ]
    })
  }

  return (
    <aside className="w-72 bg-white border-r border-gray-200 flex flex-col h-screen sticky top-0">
      {/* Header */}
      <div className="p-6 border-b border-gray-200">
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 bg-gradient-to-br from-primary-600 to-primary-700 rounded-lg flex items-center justify-center shadow-lg">
            <Package className="h-6 w-6 text-white" />
          </div>
          <div>
            <h1 className="text-lg font-bold text-gray-900">Daily Cash Memo</h1>
            <p className="text-xs text-gray-500">Management System</p>
          </div>
        </div>
      </div>

      {/* Navigation */}
      <nav className="flex-1 overflow-y-auto p-4 space-y-6 sidebar-scrollbar">
        {menuGroups.map((group, groupIndex) => (
          <div key={groupIndex} className="space-y-1">
            <h3 className="px-3 text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">
              {group.title}
            </h3>
            {group.items.map((item) => {
              const Icon = item.icon
              return (
                <NavLink
                  key={item.path}
                  to={item.path}
                  className={({ isActive }) =>
                    cn(
                      "flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-all duration-200 group relative",
                      isActive
                        ? 'bg-primary-50 text-primary-700 font-semibold shadow-sm'
                        : 'text-gray-700 hover:bg-gray-50 hover:text-gray-900'
                    )
                  }
                >
                  {({ isActive }) => (
                    <>
                      <div className={cn(
                        "flex items-center justify-center w-8 h-8 rounded-md transition-all",
                        isActive 
                          ? "bg-primary-600 text-white shadow-md" 
                          : "bg-gray-100 text-gray-600 group-hover:bg-gray-200"
                      )}>
                        <Icon className="h-4 w-4" />
                      </div>
                      <span className="text-sm font-medium flex-1">{item.label}</span>
                      {isActive && (
                        <div className="absolute right-0 top-1/2 -translate-y-1/2 w-1 h-6 bg-primary-600 rounded-l-full"></div>
                      )}
                    </>
                  )}
                </NavLink>
              )
            })}
          </div>
        ))}
      </nav>

      {/* User Info & Logout */}
      <div className="p-4 border-t border-gray-200 bg-gray-50 space-y-3">
        <div className="px-3 py-2 rounded-lg bg-white border border-gray-200">
          <p className="text-xs font-semibold text-gray-500 uppercase mb-1">Logged in as</p>
          <p className="text-sm font-semibold text-gray-900">{user?.name || 'User'}</p>
          <p className="text-xs text-gray-500 mt-0.5">{user?.email}</p>
        </div>
        <Button
          onClick={logout}
          variant="outline"
          className="w-full justify-center"
          size="sm"
        >
          <LogOut className="h-4 w-4 mr-2" />
          Logout
        </Button>
      </div>
    </aside>
  )
}

export default Sidebar
