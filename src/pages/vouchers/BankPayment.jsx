import { useState } from 'react'
import { ArrowLeft, Save, FileText } from 'lucide-react'
import { useNavigate } from 'react-router-dom'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Select from '../../components/ui/Select'
import { useForm } from 'react-hook-form'

const BankPayment = () => {
  const navigate = useNavigate()
  const { register, handleSubmit, formState: { errors }, watch } = useForm()
  const [entries, setEntries] = useState([
    { account: '', amount: '', description: '', type: 'debit' },
    { account: '', amount: '', description: '', type: 'credit' },
  ])

  const accounts = [
    { value: '1002', label: 'Bank - HBL Main' },
    { value: '1003', label: 'Bank - UBL Savings' },
    { value: '4001', label: 'Supplier ABC' },
    { value: '5001', label: 'Expense - Mazdoor' },
    { value: '5002', label: 'Expense - Electricity' },
  ]

  const onSubmit = (data) => {
    console.log('Bank Payment Voucher:', data)
    // Handle submission
  }

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
          <h1 className="text-3xl font-bold text-gray-900">Bank Payment Voucher</h1>
          <p className="text-gray-600 mt-1">Record payments made from bank account</p>
        </div>
      </div>

      <Card>
        <form onSubmit={handleSubmit(onSubmit)}>
          <div className="grid grid-cols-2 gap-4 mb-6">
            <Input
              label="Voucher Number"
              name="voucherNumber"
              defaultValue="BP-20240115-001"
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
            <Select
              label="Bank Account"
              name="bankAccount"
              register={register}
              required
              options={accounts.filter(a => a.label.includes('Bank'))}
              placeholder="Select bank account"
            />
          </div>

          <div className="mb-6">
            <h3 className="text-lg font-semibold mb-4">Voucher Entries</h3>
            <div className="space-y-4">
              <div className="grid grid-cols-12 gap-4 items-end">
                <div className="col-span-4">
                  <Select
                    label="Debit Account"
                    name="debitAccount"
                    register={register}
                    required
                    options={accounts}
                    placeholder="Select account"
                  />
                </div>
                <div className="col-span-3">
                  <Input
                    label="Amount"
                    name="debitAmount"
                    type="number"
                    register={register}
                    required
                    placeholder="0.00"
                  />
                </div>
                <div className="col-span-5">
                  <Input
                    label="Description"
                    name="debitDescription"
                    register={register}
                    placeholder="Payment description"
                  />
                </div>
              </div>
              <div className="grid grid-cols-12 gap-4 items-end">
                <div className="col-span-4">
                  <Select
                    label="Credit Account"
                    name="creditAccount"
                    register={register}
                    required
                    options={accounts.filter(a => a.label.includes('Bank'))}
                    placeholder="Select bank account"
                  />
                </div>
                <div className="col-span-3">
                  <Input
                    label="Amount"
                    name="creditAmount"
                    type="number"
                    register={register}
                    required
                    placeholder="0.00"
                  />
                </div>
                <div className="col-span-5">
                  <Input
                    label="Description"
                    name="creditDescription"
                    register={register}
                    placeholder="Payment description"
                  />
                </div>
              </div>
            </div>
            <div className="mt-4 p-4 bg-gray-50 rounded-lg">
              <div className="flex justify-between items-center">
                <span className="font-medium">Total Amount:</span>
                <span className="text-xl font-bold">PKR 0.00</span>
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
            <Button type="submit" variant="primary">
              <Save className="h-5 w-5 mr-2" />
              Post Voucher
            </Button>
          </div>
        </form>
      </Card>
    </div>
  )
}

export default BankPayment

