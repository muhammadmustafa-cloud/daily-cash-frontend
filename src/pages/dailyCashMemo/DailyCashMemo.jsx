import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import toast from 'react-hot-toast'
import { Plus, Trash2, Calendar, Save, FileText, Eye, X, Upload, Download, FileSpreadsheet } from 'lucide-react'
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
import api from '../../config/api'
import { formatCurrency, formatDate } from '../../utils/formatters'
import { exportToPDF, exportToExcel } from '../../utils/exportUtils'

const DailyCashMemo = () => {
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0])
  const [memo, setMemo] = useState(null)
  const [loading, setLoading] = useState(false)
  const [previousBalance, setPreviousBalance] = useState(0)
  const [isEntryModalOpen, setIsEntryModalOpen] = useState(false)
  const [editingEntry, setEditingEntry] = useState(null)
  const [entryType, setEntryType] = useState('credit') // 'credit' or 'debit'
  const [isNotesModalOpen, setIsNotesModalOpen] = useState(false)
  const [imagePreview, setImagePreview] = useState(null)
  const [selectedImage, setSelectedImage] = useState(null)
  const [isImageModalOpen, setIsImageModalOpen] = useState(false)

  const { register, handleSubmit, formState: { errors }, reset, setValue, watch, setError, clearErrors } = useForm({
    mode: 'onChange',
    defaultValues: {
      name: '',
      description: '',
      amount: '',
      notes: ''
    }
  })
  const notes = watch('notes')

  useEffect(() => {
    fetchMemo()
    fetchPreviousBalance()
  }, [selectedDate])

  const fetchPreviousBalance = async () => {
    try {
      const response = await api.get('/daily-cash-memos/previous-balance', {
        params: { date: selectedDate }
      })
      setPreviousBalance(response.data.data.previousBalance || 0)
    } catch (error) {
      console.error('Error fetching previous balance:', error)
    }
  }

  const fetchMemo = async () => {
    try {
      setLoading(true)
      const response = await api.get(`/daily-cash-memos/date/${selectedDate}`)
      setMemo(response.data.data.memo)
      setValue('notes', response.data.data.memo.notes || '')
    } catch (error) {
      if (error.response?.status === 404) {
        setMemo(null)
        setValue('notes', '')
      } else {
        toast.error('Failed to fetch daily cash memo')
      }
    } finally {
      setLoading(false)
    }
  }

  const calculateTotalCredit = () => {
    if (!memo) return previousBalance
    const entriesTotal = (memo.creditEntries || []).reduce((sum, entry) => sum + (entry.amount || 0), 0)
    return previousBalance + entriesTotal
  }

  const calculateTotalDebit = () => {
    if (!memo) return 0
    return (memo.debitEntries || []).reduce((sum, entry) => sum + (entry.amount || 0), 0)
  }

  const calculateClosingBalance = () => {
    return calculateTotalCredit() - calculateTotalDebit()
  }

  const handleAddEntry = (type) => {
    setEntryType(type)
    setEditingEntry(null)
    reset({
      name: '',
      description: '',
      amount: '',
      notes: memo?.notes || ''
    }, {
      keepErrors: false,
      keepDirty: false,
      keepIsSubmitted: false,
      keepTouched: false,
      keepIsValid: false,
      keepSubmitCount: false
    })
    setImagePreview(null)
    setIsEntryModalOpen(true)
  }

  const handleEditEntry = (entry, type) => {
    setEntryType(type)
    setEditingEntry(entry)
    reset({
      name: entry.name,
      description: entry.description || '',
      amount: entry.amount,
      notes: memo?.notes || ''
    }, {
      keepErrors: false,
      keepDirty: false,
      keepIsSubmitted: false,
      keepTouched: false,
      keepIsValid: false,
      keepSubmitCount: false
    })
    setImagePreview(entry.image || null)
    setIsEntryModalOpen(true)
  }

  const handleDeleteEntry = async (entry, type) => {
    if (!window.confirm('Are you sure you want to delete this entry?')) {
      return
    }

    try {
      const updatedMemo = { ...memo }
      if (type === 'credit') {
        updatedMemo.creditEntries = updatedMemo.creditEntries.filter(e => 
          e._id ? e._id !== entry._id : e !== entry
        )
      } else {
        updatedMemo.debitEntries = updatedMemo.debitEntries.filter(e => 
          e._id ? e._id !== entry._id : e !== entry
        )
      }

      if (memo?._id) {
        await api.put(`/daily-cash-memos/${memo._id}`, {
          creditEntries: updatedMemo.creditEntries,
          debitEntries: updatedMemo.debitEntries,
          previousBalance: previousBalance
        })
        toast.success('Entry deleted successfully')
      }
      fetchMemo()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to delete entry')
    }
  }

  const handleImageChange = (e) => {
    const file = e.target.files[0]
    if (file) {
      // Validate file type
      if (!file.type.startsWith('image/')) {
        toast.error('Please select an image file')
        return
      }
      // Validate file size (max 5MB)
      if (file.size > 5 * 1024 * 1024) {
        toast.error('Image size should be less than 5MB')
        return
      }
      // Convert to base64
      const reader = new FileReader()
      reader.onloadend = () => {
        setImagePreview(reader.result)
      }
      reader.readAsDataURL(file)
    }
  }

  const removeImage = () => {
    setImagePreview(null)
  }

  const validateEntry = (data) => {
    const errors = {};
    
    if (!data.name || data.name.trim() === '') {
      errors.name = 'Name is required';
    }
    
    if (!data.amount || isNaN(parseFloat(data.amount))) {
      errors.amount = 'Valid amount is required';
    } else if (parseFloat(data.amount) <= 0) {
      errors.amount = 'Amount must be greater than zero';
    }
    
    // Validate date is not in the future
    const selected = new Date(selectedDate);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    if (selected > today) {
      errors.date = 'Cannot select a future date';
    }
    
    return Object.keys(errors).length === 0 ? null : errors;
  };

  const onEntrySubmit = async (data) => {
    // Clear any existing errors
    clearErrors();
    
    // Manual validation
    const validationErrors = {};
    
    if (!data.name || data.name.trim() === '') {
      validationErrors.name = { message: 'Name is required' };
    } else if (data.name.length > 100) {
      validationErrors.name = { message: 'Name must be less than 100 characters' };
    }
    
    if (data.description && data.description.length > 500) {
      validationErrors.description = { message: 'Description must be less than 500 characters' };
    }
    
    if (!data.amount || isNaN(parseFloat(data.amount)) || parseFloat(data.amount) <= 0) {
      validationErrors.amount = { message: 'Please enter a valid amount greater than zero' };
    }
    
    if (Object.keys(validationErrors).length > 0) {
      // Set the errors in the form
      Object.entries(validationErrors).forEach(([field, error]) => {
        setError(field, { type: 'manual', message: error.message });
      });
      return;
    }

    try {
      const entry = {
        name: data.name.trim(),
        description: (data.description || '').trim(),
        amount: parseFloat(parseFloat(data.amount).toFixed(2)), // Ensure 2 decimal places
        image: imagePreview || undefined
      };

      let updatedMemo = memo ? { ...memo } : {
        date: selectedDate,
        creditEntries: [],
        debitEntries: [],
        previousBalance: previousBalance
      }

      if (entryType === 'credit') {
        if (editingEntry) {
          updatedMemo.creditEntries = updatedMemo.creditEntries.map(e =>
            (e._id && editingEntry._id && e._id === editingEntry._id) || e === editingEntry ? entry : e
          )
        } else {
          updatedMemo.creditEntries = [...(updatedMemo.creditEntries || []), entry]
        }
      } else {
        if (editingEntry) {
          updatedMemo.debitEntries = updatedMemo.debitEntries.map(e =>
            (e._id && editingEntry._id && e._id === editingEntry._id) || e === editingEntry ? entry : e
          )
        } else {
          updatedMemo.debitEntries = [...(updatedMemo.debitEntries || []), entry]
        }
      }

      if (memo?._id) {
        await api.put(`/daily-cash-memos/${memo._id}`, {
          creditEntries: updatedMemo.creditEntries,
          debitEntries: updatedMemo.debitEntries,
          previousBalance: previousBalance
        })
        toast.success('Entry updated successfully')
      } else {
        const response = await api.post('/daily-cash-memos', {
          date: selectedDate,
          creditEntries: updatedMemo.creditEntries,
          debitEntries: updatedMemo.debitEntries,
          previousBalance: previousBalance
        })
        setMemo(response.data.data.memo)
        toast.success('Entry added successfully')
      }

      setIsEntryModalOpen(false)
      reset()
      setImagePreview(null)
      fetchMemo()
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save entry')
    }
  }

  const handleViewImage = (image) => {
    setSelectedImage(image)
    setIsImageModalOpen(true)
  }

  const handleSaveNotes = async (data) => {
    try {
      if (memo?._id) {
        await api.put(`/daily-cash-memos/${memo._id}`, {
          notes: data.notes || '',
          creditEntries: memo.creditEntries,
          debitEntries: memo.debitEntries,
          previousBalance: previousBalance
        })
      } else {
        await api.post('/daily-cash-memos', {
          date: selectedDate,
          creditEntries: memo?.creditEntries || [],
          debitEntries: memo?.debitEntries || [],
          previousBalance: previousBalance,
          notes: data.notes || ''
        })
        fetchMemo()
      }
      setIsNotesModalOpen(false)
      toast.success('Notes saved successfully')
    } catch (error) {
      toast.error(error.response?.data?.message || 'Failed to save notes')
    }
  }

  const getDayName = (dateString) => {
    const date = new Date(dateString)
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    return days[date.getDay()]
  }

  const handleExportPDF = () => {
    try {
      exportToPDF(memo, selectedDate, previousBalance)
      toast.success('PDF exported successfully')
    } catch (error) {
      console.error('Error exporting PDF:', error)
      toast.error(error.message || 'Failed to export PDF')
    }
  }

  const handleExportExcel = () => {
    try {
      exportToExcel(memo, selectedDate, previousBalance)
      toast.success('Excel exported successfully')
    } catch (error) {
      console.error('Error exporting Excel:', error)
      toast.error('Failed to export Excel')
    }
  }

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900">Daily Cash Memo</h1>
          <p className="text-gray-600 mt-1.5">Manage daily cash in and cash out transactions</p>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <Calendar className="h-5 w-5 text-gray-500" />
            <input
              type="date"
              value={selectedDate}
              onChange={(e) => {
                setSelectedDate(e.target.value)
                setMemo(null)
              }}
              className="px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
            />
          </div>
          <Button
            variant="outline"
            onClick={handleExportPDF}
            className="bg-red-50 hover:bg-red-100 text-red-700 border-red-300"
            disabled={!memo}
          >
            <Download className="h-4 w-4 mr-2" />
            Download PDF
          </Button>
          <Button
            variant="outline"
            onClick={handleExportExcel}
            className="bg-green-50 hover:bg-green-100 text-green-700 border-green-300"
            disabled={!memo}
          >
            <FileSpreadsheet className="h-4 w-4 mr-2" />
            Download Excel
          </Button>
          {memo && (
            <Button
              variant="outline"
              onClick={() => setIsNotesModalOpen(true)}
            >
              <FileText className="h-4 w-4 mr-2" />
              Notes
            </Button>
          )}
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center h-64">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-primary-600"></div>
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* CREDIT (Cash In) Section */}
          <Card>
            <CardHeader>
              <div className="flex justify-between items-center">
                <div>
                  <CardTitle className="text-green-700">CREDIT (Cash In)</CardTitle>
                  <CardDescription>All incoming cash transactions</CardDescription>
                </div>
                <Button
                  size="sm"
                  onClick={() => handleAddEntry('credit')}
                  className="bg-green-600 hover:bg-green-700"
                >
                  <Plus className="h-4 w-4 mr-2" />
                  Add Entry
                </Button>
              </div>
            </CardHeader>
            <CardContent>
              <div className="space-y-4">
                {/* Previous Balance */}
                <div className="bg-gray-50 p-4 rounded-lg border border-gray-200">
                  <div className="flex justify-between items-center">
                    <div>
                      <p className="font-semibold text-gray-700">Previous Balance</p>
                      <p className="text-sm text-gray-500">Carried forward from previous day</p>
                    </div>
                    <p className="text-xl font-bold text-gray-900">
                      {formatCurrency(previousBalance)}
                    </p>
                  </div>
                </div>

                {/* Credit Entries */}
                <div className="border border-gray-200 rounded-lg overflow-hidden">
                  <table className="w-full">
                    <thead className="bg-gray-50">
                      <tr>
                        <th className="px-4 py-3 text-left text-xs font-semibold text-gray-700 uppercase">Name</th>
                        <th className="px-4 py-3 text-left text-xs font-semibold text-gray-700 uppercase">Description</th>
                        <th className="px-4 py-3 text-right text-xs font-semibold text-gray-700 uppercase">Amount</th>
                        <th className="px-4 py-3 text-center text-xs font-semibold text-gray-700 uppercase">Image</th>
                        <th className="px-4 py-3 text-center text-xs font-semibold text-gray-700 uppercase">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200">
                      {memo?.creditEntries?.map((entry, index) => (
                        <tr key={entry._id || index} className="hover:bg-gray-50">
                          <td className="px-4 py-3 text-sm text-gray-900">{entry.name}</td>
                          <td className="px-4 py-3 text-sm text-gray-600">{entry.description || '-'}</td>
                          <td className="px-4 py-3 text-sm text-right font-semibold text-gray-900">
                            {formatCurrency(entry.amount)}
                          </td>
                          <td className="px-4 py-3 text-center">
                            {entry.image ? (
                              <button
                                onClick={() => handleViewImage(entry.image)}
                                className="text-blue-600 hover:text-blue-800 transition-colors"
                                title="View image"
                                aria-label={`View image for ${entry.name}`}
                              >
                                <Eye className="h-5 w-5" />
                              </button>
                            ) : (
                              <span className="text-gray-400">-</span>
                            )}
                          </td>
                          <td className="px-4 py-3 text-center">
                            <div className="flex justify-center gap-2">
                              <button
                                onClick={() => handleEditEntry(entry, 'credit')}
                                className="text-blue-600 hover:text-blue-800 text-sm"
                                aria-label={`Edit ${entry.name}`}
                              >
                                Edit
                              </button>
                              <button
                                onClick={() => handleDeleteEntry(entry, 'credit')}
                                className="text-red-600 hover:text-red-800 text-sm"
                                aria-label={`Delete ${entry.name}`}
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                      {(!memo?.creditEntries || memo.creditEntries.length === 0) && (
                        <tr>
                          <td colSpan="5" className="px-4 py-8 text-center text-gray-500">
                            No credit entries yet. Click "Add Entry" to add one.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Total Credit */}
                <div className="bg-green-50 p-4 rounded-lg border-2 border-green-200">
                  <div className="flex justify-between items-center">
                    <p className="text-lg font-bold text-green-800">Total</p>
                    <p className="text-2xl font-bold text-green-900">
                      {formatCurrency(calculateTotalCredit())}
                    </p>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>

          {/* DEBIT (Cash Out) Section */}
          <Card>
            <CardHeader>
              <div className="flex justify-between items-center">
                <div>
                  <CardTitle className="text-red-700">DEBIT (Cash Out)</CardTitle>
                  <CardDescription>All outgoing cash transactions</CardDescription>
                </div>
                <Button
                  size="sm"
                  onClick={() => handleAddEntry('debit')}
                  className="bg-red-600 hover:bg-red-700"
                >
                  <Plus className="h-4 w-4 mr-2" />
                  Add Entry
                </Button>
              </div>
            </CardHeader>
            <CardContent>
              <div className="space-y-4">
                {/* Debit Entries */}
                <div className="border border-gray-200 rounded-lg overflow-hidden">
                  <table className="w-full">
                    <thead className="bg-gray-50">
                      <tr>
                        <th className="px-4 py-3 text-left text-xs font-semibold text-gray-700 uppercase">Name</th>
                        <th className="px-4 py-3 text-left text-xs font-semibold text-gray-700 uppercase">Description</th>
                        <th className="px-4 py-3 text-right text-xs font-semibold text-gray-700 uppercase">Amount</th>
                        <th className="px-4 py-3 text-center text-xs font-semibold text-gray-700 uppercase">Image</th>
                        <th className="px-4 py-3 text-center text-xs font-semibold text-gray-700 uppercase">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200">
                      {memo?.debitEntries?.map((entry, index) => (
                        <tr key={entry._id || index} className="hover:bg-gray-50">
                          <td className="px-4 py-3 text-sm text-gray-900">{entry.name}</td>
                          <td className="px-4 py-3 text-sm text-gray-600">{entry.description || '-'}</td>
                          <td className="px-4 py-3 text-sm text-right font-semibold text-gray-900">
                            {formatCurrency(entry.amount)}
                          </td>
                          <td className="px-4 py-3 text-center">
                            {entry.image ? (
                              <button
                                onClick={() => handleViewImage(entry.image)}
                                className="text-blue-600 hover:text-blue-800 transition-colors"
                                title="View image"
                                aria-label={`View image for ${entry.name}`}
                              >
                                <Eye className="h-5 w-5" />
                              </button>
                            ) : (
                              <span className="text-gray-400">-</span>
                            )}
                          </td>
                          <td className="px-4 py-3 text-center">
                            <div className="flex justify-center gap-2">
                              <button
                                onClick={() => handleEditEntry(entry, 'debit')}
                                className="text-blue-600 hover:text-blue-800 text-sm"
                                aria-label={`Edit ${entry.name}`}
                              >
                                Edit
                              </button>
                              <button
                                onClick={() => handleDeleteEntry(entry, 'debit')}
                                className="text-red-600 hover:text-red-800 text-sm"
                                aria-label={`Delete ${entry.name}`}
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                      {(!memo?.debitEntries || memo.debitEntries.length === 0) && (
                        <tr>
                          <td colSpan="5" className="px-4 py-8 text-center text-gray-500">
                            No debit entries yet. Click "Add Entry" to add one.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Total Debit */}
                <div className="bg-red-50 p-4 rounded-lg border-2 border-red-200">
                  <div className="flex justify-between items-center">
                    <p className="text-lg font-bold text-red-800">Total</p>
                    <p className="text-2xl font-bold text-red-900">
                      {formatCurrency(calculateTotalDebit())}
                    </p>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      )}

      {/* Closing Balance Card */}
      <Card className="bg-gradient-to-r from-blue-50 to-indigo-50 border-2 border-blue-200">
        <CardContent className="pt-6">
          <div className="flex justify-between items-center">
            <div>
              <p className="text-sm font-semibold text-gray-600 uppercase tracking-wide">
                {getDayName(selectedDate)} - {formatDate(selectedDate)}
              </p>
              <p className="text-2xl font-bold text-gray-900 mt-2">Closing Balance</p>
            </div>
            <div className="text-right">
              <p className="text-sm text-gray-600 mb-1">Total Credit - Total Debit</p>
              <p className="text-4xl font-bold text-blue-900">
                {formatCurrency(calculateClosingBalance())}
              </p>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Entry Modal */}
      <Dialog open={isEntryModalOpen} onOpenChange={setIsEntryModalOpen}>
        <DialogContent className="sm:max-w-[500px]">
          <DialogHeader>
            <DialogTitle>
              {editingEntry ? 'Edit' : 'Add'} {entryType === 'credit' ? 'Credit' : 'Debit'} Entry
            </DialogTitle>
            <DialogDescription>
              {editingEntry ? 'Update the entry details below' : 'Add a new entry with the details below'}
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(onEntrySubmit)} className="space-y-4">
            <FormInput
              label="Name"
              id="name"
              {...register('name', {
                required: 'Name is required',
                maxLength: {
                  value: 100,
                  message: 'Name must be less than 100 characters',
                },
                validate: (value) => {
                  const trimmed = value.trim();
                  return trimmed.length > 0 || 'Name cannot be just whitespace';
                },
              })}
              error={errors.name}
              placeholder="Enter name"
              aria-required="true"
              aria-invalid={errors.name ? 'true' : 'false'}
              aria-describedby={errors.name ? 'name-error' : undefined}
            />

            <FormInput
              label="Description"
              id="description"
              {...register('description', {
                maxLength: {
                  value: 500,
                  message: 'Description must be less than 500 characters',
                },
              })}
              error={errors.description}
              placeholder="Enter description (optional)"
              as="textarea"
              rows={3}
            />

            <FormInput
              label="Amount"
              id="amount"
              type="number"
              step="0.01"
              min="0.01"
              {...register('amount', {
                required: 'Amount is required',
                min: {
                  value: 0.01,
                  message: 'Amount must be greater than zero',
                },
                max: {
                  value: 1000000,
                  message: 'Amount is too large',
                },
                valueAsNumber: true,
              })}
              error={errors.amount}
              placeholder="0.00"
              inputMode="decimal"
              aria-required="true"
              aria-invalid={errors.amount ? 'true' : 'false'}
              aria-describedby={errors.amount ? 'amount-error' : undefined}
            />

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Image (Optional)
              </label>
              {imagePreview && (
                <div className="relative inline-block mb-2">
                  <img
                    src={imagePreview}
                    alt="Preview"
                    className="h-32 w-32 object-cover rounded-lg border border-gray-300"
                  />
                  <button
                    type="button"
                    onClick={removeImage}
                    className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full p-1 hover:bg-red-600 transition-colors"
                    aria-label="Remove image"
                  >
                    <X className="h-4 w-4" />
                  </button>
                </div>
              )}
              <div>
                <label className="flex items-center justify-center w-full px-4 py-2 border border-gray-300 rounded-lg cursor-pointer hover:bg-gray-50 transition-colors">
                  <Upload className="h-5 w-5 mr-2 text-gray-500" />
                  <span className="text-sm text-gray-700">
                    {imagePreview ? 'Change Image' : 'Upload Image'}
                  </span>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={handleImageChange}
                    className="hidden"
                  />
                </label>
                <p className="text-xs text-gray-500 mt-1">Max size: 5MB</p>
              </div>
            </div>
            <DialogFooter>
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  setIsEntryModalOpen(false)
                  reset()
                  setImagePreview(null)
                }}
              >
                Cancel
              </Button>
              <Button type="submit">
                {editingEntry ? 'Update' : 'Add'} Entry
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* Notes Modal */}
      <Dialog open={isNotesModalOpen} onOpenChange={setIsNotesModalOpen}>
        <DialogContent className="sm:max-w-[500px]">
          <DialogHeader>
            <DialogTitle>Notes</DialogTitle>
            <DialogDescription>
              Add any additional notes for this daily cash memo.
            </DialogDescription>
          </DialogHeader>
          <form onSubmit={handleSubmit(handleSaveNotes)} className="space-y-4">
            <div>
              <label htmlFor="notes" className="block text-sm font-medium text-gray-700 mb-2">
                Notes
              </label>
              <textarea
                id="notes"
                {...register('notes')}
                rows={6}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500"
                placeholder="Enter any additional notes..."
                aria-describedby="notes-description"
              />
              <p id="notes-description" className="mt-1 text-sm text-gray-500">
                Add any notes or comments about this day's transactions.
              </p>
            </div>
            <DialogFooter>
              <Button
                type="button"
                variant="outline"
                onClick={() => setIsNotesModalOpen(false)}
                aria-label="Cancel and close notes dialog"
              >
                Cancel
              </Button>
              <Button type="submit" aria-label="Save notes">
                <Save className="h-4 w-4 mr-2" />
                Save Notes
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* Image View Modal */}
      <Dialog open={isImageModalOpen} onOpenChange={setIsImageModalOpen}>
        <DialogContent className="sm:max-w-[600px]">
          <DialogHeader>
            <DialogTitle>Entry Image</DialogTitle>
            <DialogDescription>
              View the uploaded image for this entry
            </DialogDescription>
          </DialogHeader>
          <div className="flex justify-center items-center p-4">
            {selectedImage && (
              <img
                src={selectedImage}
                alt="Entry receipt or document"
                className="max-w-full max-h-[500px] rounded-lg border border-gray-300 object-contain"
              />
            )}
          </div>
          <DialogFooter>
            <Button
              type="button"
              variant="outline"
              onClick={() => setIsImageModalOpen(false)}
              aria-label="Close image viewer"
            >
              Close
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  )
}

export default DailyCashMemo

