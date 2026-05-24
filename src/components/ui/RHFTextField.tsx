import { TextField } from '@mui/material'
import { FC } from 'react'
import { Controller } from 'react-hook-form'

interface ITextField {
   type?: string
   size?: 'small' | 'medium'
   name: string
   label?: string
   control: any
   multiline?: boolean
   rows?: number
   rules?: { required: string }
   disabled?: boolean
   required?: boolean
}

const RHFTextField: FC<ITextField> = ({ type = 'text', size = 'medium', name, label = '', control, multiline = false, rows = 1, rules = {}, disabled = false, required = false }) => (
   <Controller
      name={name}
      control={control}
      defaultValue=""
      rules={required ? rules : {}}
      render={({ field, fieldState: { error } }) => (
         <TextField
            type={type}
            size={size}
            label={label}
            variant="outlined"
            fullWidth
            multiline={multiline}
            rows={rows}
            error={!!error}
            helperText={error ? error.message : ''}
            {...field}
            disabled={disabled}
         />
      )}
   />
)

export default RHFTextField
