import { ArrowLeft, Save, FileText } from 'lucide-react'
import { useNavigate } from 'react-router-dom'
import Card from '../../components/ui/Card'
import Button from '../../components/ui/Button'
import Input from '../../components/ui/Input'
import Select from '../../components/ui/Select'
import { useForm } from 'react-hook-form'

const BankTransfer = () => {
  const navigate = useNavigate()
  const { register, handleSubmit } = useForm()

  const bankAccounts = [
    { value: '1002', label: 'Bank - HBL Main' },
    { value: '1003', label: 'Bank - UBL Savings' },
  ]

  const onSubmit = (data) => {
    console.log('Bank Transfer Voucher:', data)
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
          <h1 className="text-3xl font-bold text-gray-900">Bank Account Transfer</h1>
          <p className="text-gray-600 mt-1">Transfer money between bank accounts</p>
        </div>
      </div>

      <Card>
        <form onSubmit={handleSubmit(onSubmit)}>
          <div className="grid grid-cols-2 gap-4 mb-6">
            <Input
              label="Voucher Number"
              name="voucherNumber"
              defaultValue="BT-20240119-001"
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
            <h3 className="text-lg font-semibold mb-4">Transfer Details</h3>
            <div className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <Select
                  label="From Bank Account"
                  name="fromAccount"
                  register={register}
                  required
                  options={bankAccounts}
                  placeholder="Select source account"
                />
                <Select
                  label="To Bank Account"
                  name="toAccount"
                  register={register}
                  required
                  options={bankAccounts}
                  placeholder="Select destination account"
                />
              </div>
              <Input
                label="Transfer Amount"
                name="amount"
                type="number"
                register={register}
                required
                placeholder="0.00"
              />
              <Input
                label="Description"
                name="description"
                register={register}
                placeholder="Transfer description"
              />
            </div>
            <div className="mt-4 p-4 bg-gray-50 rounded-lg">
              <div className="flex justify-between items-center">
                <span className="font-medium">Transfer Amount:</span>
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
              Post Transfer
            </Button>
          </div>
        </form>
      </Card>
    </div>
  )
}

export default BankTransfer

