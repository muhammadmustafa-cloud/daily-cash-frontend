import * as React from "react"
import { cn } from "../../lib/utils"

const Input = React.forwardRef(({ className, type = "text", label, name, id, required, error, placeholder, register, ...props }, ref) => {
  const inputId = id || name;
  
  // If register is provided, use it to register the input
  const registerProps = register ? register(name) : {};
  return (
    <div className="space-y-2">
      {label && (
        <label htmlFor={inputId} className="block text-sm font-semibold text-gray-700">
          {label}
          {required && <span className="text-red-500 ml-1">*</span>}
        </label>
      )}
      <input
        type={type}
        id={inputId}
        name={name}
        className={cn(
          "flex h-11 w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm",
          "placeholder:text-gray-400",
          "focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary-500 focus-visible:ring-offset-2",
          "disabled:cursor-not-allowed disabled:opacity-50",
          "transition-all duration-200",
          "hover:border-gray-400",
          error && "border-red-500 focus-visible:ring-red-500",
          className
        )}
        ref={ref}
        placeholder={placeholder}
        {...registerProps}
        {...props}
      />
      {error && (
        <p className="text-sm text-red-600 font-medium">{error}</p>
      )}
    </div>
  )}
)
Input.displayName = "Input"

export default Input
