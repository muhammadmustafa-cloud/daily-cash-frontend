import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom'
import { Toaster } from 'react-hot-toast'
import { AuthProvider } from './contexts/AuthContext'
import ProtectedRoute from './components/auth/ProtectedRoute'
import Layout from './components/layout/Layout'
import Login from './pages/auth/Login'
import Dashboard from './pages/Dashboard'
import Customers from './pages/customers/Customers'
import Suppliers from './pages/suppliers/Suppliers'
import Items from './pages/items/Items'
import Inventory from './pages/inventory/Inventory'
import Expenses from './pages/expenses/Expenses'
import Mazdoors from './pages/mazdoors/Mazdoors'
import Transactions from './pages/transactions/Transactions'
import Accounts from './pages/accounts/Accounts'
import Payments from './pages/vouchers/Payments'
import Users from './pages/users/Users'
import DailyCashMemo from './pages/dailyCashMemo/DailyCashMemo'
import LabourRates from './pages/labourRates/LabourRates'
import LabourExpenses from './pages/labourExpenses/LabourExpenses'

function App() {
  return (
    <AuthProvider>
      <Router>
        <Toaster position="top-right" />
        <Routes>
          <Route path="/login" element={<Login />} />
          <Route
            path="/"
            element={
              <ProtectedRoute>
                <Layout />
              </ProtectedRoute>
            }
          >
            <Route index element={<Dashboard />} />
            <Route path="customers" element={<Customers />} />
            <Route path="suppliers" element={<Suppliers />} />
            <Route path="items" element={<Items />} />
            <Route path="inventory" element={<Inventory />} />
            <Route path="expenses" element={<Expenses />} />
            <Route path="mazdoors" element={<Mazdoors />} />
            <Route path="transactions" element={<Transactions />} />
            <Route path="accounts" element={<Accounts />} />
            <Route path="payments" element={<Payments />} />
            <Route path="daily-cash-memo" element={<DailyCashMemo />} />
            <Route path="labour" element={<LabourRates />} />
            <Route path="labour-expenses" element={<LabourExpenses />} />
            <Route path="users" element={<Users />} />
          </Route>
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </Router>
    </AuthProvider>
  )
}

export default App

