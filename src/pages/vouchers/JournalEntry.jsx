import { useState } from 'react'
import { ArrowLeft, Save, FileText, Plus, Trash2 } from 'lucide-react'
import { useNavigate } from 'react-router-dom'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Select from '../../components/ui/Select'
import { useForm } from 'react-hook-form'
import { formatCurrency } from '../../utils/formatters'

const JournalEntry = () => {
  const navigate = useNavigate()
  const { register, handleSubmit } = useForm()
  const [entries, setEntries] = useState([
    { account: '', amount: '', description: '', type: 'debit' },
    { account: '', amount: '', description: '', type: 'credit' },
  ])

  const accounts = [
    { value: '1001', label: 'Cash Account' },
    { value: '1002', label: 'Bank - HBL Main' },
    { value: '2001', label: 'Accounts Payable' },
    { value: '3001', label: 'Accounts Receivable' },
    { value: '4001', label: 'Sales Income' },
    { value: '5001', label: 'Expense - Mazdoor' },
  ]

  const addEntry = () => {
    setEntries([...entries, { account: '', amount: '', description: '', type: 'debit' }])
  }

  const removeEntry = (index) => {
    setEntries(entries.filter((_, i) => i !== index))
  }

  const onSubmit = (data) => {
    console.log('Journal Entry:', data, entries)
  }

  const totalDebit = entries
    .filter(e => e.type === 'debit')
    .reduce((sum, e) => sum + (parseFloat(e.amount) || 0), 0)
  const totalCredit = entries
    .filter(e => e.type === 'credit')
    .reduce((sum, e) => sum + (parseFloat(e.amount) || 0), 0)

  return (
    <div>
      <div className="flex items-center gap-4 mb-6">
        <Button
          variant="secondary"
          onClick={() => navigate('/vouchers')}
        >
          <ArrowLeft className="h-5 w-5 mr-2" />
          Back
        </Button>
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Journal Entry</h1>
          <p className="text-gray-600 mt-1">General accounting entries (adjustments, corrections)</p>
        </div>
      </div>

      <Card>
        <form onSubmit={handleSubmit(onSubmit)}>
          <div className="grid grid-cols-2 gap-4 mb-6">
            <Input
              label="Voucher Number"
              name="voucherNumber"
              defaultValue="JE-20240120-001"
              disabled
            />
            <Input
              label="Date"
              name="date"
              type="date"
              register={register}
              required
              defaultValue={new Date().toISOString().split('T')[0]}
            />
            <Input
              label="Reference Number"
              name="referenceNumber"
              register={register}
              placeholder="Optional"
            />
          </div>

          <div className="mb-6">
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-lg font-semibold">Journal Entries</h3>
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={addEntry}
              >
                <Plus className="h-4 w-4 mr-2" />
                Add Entry
              </Button>
            </div>
            <div className="space-y-4">
              {entries.map((entry, index) => (
                <div key={index} className="grid grid-cols-12 gap-4 items-end border-b pb-4">
                  <div className="col-span-3">
                    <Select
                      label={index === 0 ? 'Account' : ''}
                      name={`entry_${index}_account`}
                      register={register}
                      options={accounts}
                      placeholder="Select account"
                    />
                  </div>
                  <div className="col-span-2">
                    <Select
                      label={index === 0 ? 'Type' : ''}
                      name={`entry_${index}_type`}
                      register={register}
                      options={[
                        { value: 'debit', label: 'Debit' },
                        { value: 'credit', label: 'Credit' },
                      ]}
                      defaultValue={entry.type}
                    />
                  </div>
                  <div className="col-span-3">
                    <Input
                      label={index === 0 ? 'Amount' : ''}
                      name={`entry_${index}_amount`}
                      type="number"
                      register={register}
                      placeholder="0.00"
                    />
                  </div>
                  <div className="col-span-3">
                    <Input
                      label={index === 0 ? 'Description' : ''}
                      name={`entry_${index}_description`}
                      register={register}
                      placeholder="Description"
                    />
                  </div>
                  <div className="col-span-1">
                    {entries.length > 2 && (
                      <Button
                        type="button"
                        variant="danger"
                        size="sm"
                        onClick={() => removeEntry(index)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    )}
                  </div>
                </div>
              ))}
            </div>
            <div className="mt-4 p-4 bg-gray-50 rounded-lg">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <span className="font-medium">Total Debit:</span>
                  <span className="ml-2 text-lg font-bold">{formatCurrency(totalDebit)}</span>
                </div>
                <div>
                  <span className="font-medium">Total Credit:</span>
                  <span className="ml-2 text-lg font-bold">{formatCurrency(totalCredit)}</span>
                </div>
              </div>
              <div className="mt-2 pt-2 border-t">
                <span className="font-medium">Difference:</span>
                <span className={`ml-2 text-lg font-bold ${
                  totalDebit === totalCredit ? 'text-green-600' : 'text-red-600'
                }`}>
                  {formatCurrency(totalDebit - totalCredit)}
                </span>
                {totalDebit === totalCredit ? (
                  <span className="ml-2 text-sm text-green-600">✓ Balanced</span>
                ) : (
                  <span className="ml-2 text-sm text-red-600">⚠ Not Balanced</span>
                )}
              </div>
            </div>
          </div>

          <div className="mb-6">
            <label className="label">Notes</label>
            <textarea
              {...register('notes')}
              rows={3}
              className="input"
              placeholder="Additional notes..."
            />
          </div>

          <div className="flex justify-end gap-3">
            <Button
              type="button"
              variant="secondary"
              onClick={() => navigate('/vouchers')}
            >
              Cancel
            </Button>
            <Button type="button" variant="outline">
              <FileText className="h-5 w-5 mr-2" />
              Save as Draft
            </Button>
            <Button
              type="submit"
              variant="primary"
              disabled={totalDebit !== totalCredit}
            >
              <Save className="h-5 w-5 mr-2" />
              Post Entry
            </Button>
          </div>
        </form>
      </Card>
    </div>
  )
}

export default JournalEntry

