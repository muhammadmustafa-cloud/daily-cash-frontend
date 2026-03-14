import * as React from "react"
import Input from "./Input"

const FormInput = React.forwardRef(
  ({ label, name, register, required, error, className, ...props }, ref) => {
    // Extract error message from error object if it exists
    const errorMessage = error?.message || error;
    
    return (
      <Input
        ref={ref}
        label={label}
        name={name}
        register={register}
        required={required}
        error={errorMessage} // Only pass the error message string
        className={className}
        {...props}
      />
    )
  }
)
FormInput.displayName = "FormInput"

export { FormInput }

