import { FormControl, FormHelperText, InputLabel, MenuItem, Select } from '@mui/material'
import { FC } from 'react'
import { Controller } from 'react-hook-form'

interface ISelect {
   name: string
   label?: string
   size?: 'small' | 'medium'
   control: any
   options: { name: string; label: string }[]
   defaultValue?: string
   rules?: { required: string }
   disabled?: boolean
   required?: boolean
}

const RHFSelect: FC<ISelect> = ({ name, label = '', size = 'medium', control, options, defaultValue = '', rules = {}, disabled = false, required = false }) => {
   return (
      <Controller
         name={name}
         control={control}
         defaultValue={defaultValue}
         rules={required ? rules : {}}
         render={({ field, fieldState: { error } }) => (
            <FormControl fullWidth size={size} disabled={disabled} error={error ? true : false}>
               <InputLabel>{label}</InputLabel>
               <Select {...field} label={label}>
                  {options.map((obj, idx) => (
                     <MenuItem key={idx} value={obj.name}>
                        {obj.label}
                     </MenuItem>
                  ))}
               </Select>
               {error && <FormHelperText error>{error.message}</FormHelperText>}
            </FormControl>
         )}
      />
   )
}

export default RHFSelect
