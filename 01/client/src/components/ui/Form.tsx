import { Replay, Search } from '@mui/icons-material'
import { Box, IconButton } from '@mui/material'
import { FC } from 'react'
import { useFormContext } from 'react-hook-form'

import RHFSelect from './RHFSelect'
import RHFTextField from './RHFTextField'

interface IDataForm {
   onSubmit: (data: { codeList: string; pg?: string; type?: string }) => void
   onReset: () => void
   hasPGField?: boolean
   hasTypeField?: boolean
}

const Form: FC<IDataForm> = ({ onSubmit, onReset, hasPGField = false, hasTypeField = false }) => {
   const { control, handleSubmit, reset } = useFormContext<{ codeList: string; pg?: string; type?: string }>()

   return (
      <Box component="form" onSubmit={handleSubmit(onSubmit)} className={`flex items-center gap-3 md:flex-row md:items-start md:gap-5 ${hasPGField || hasTypeField ? 'flex-col' : ''}`}>
         <RHFTextField size="small" name="codeList" label="Поиск..." control={control} rules={{ required: 'Обязательное поле!' }} required />

         <Box className={`flex max-w-[800px] items-start gap-2 ${hasPGField || hasTypeField ? 'w-full' : ''}`}>
            {hasPGField && (
               <RHFSelect
                  size="small"
                  name="pg"
                  label="Товарная группа"
                  control={control}
                  options={[
                     { name: 'pharma', label: 'Лекарственные средства' },
                     { name: 'water', label: 'Вода и напитки' },
                     { name: 'appliances', label: 'Бытовая техника' },
                     { name: 'tobacco', label: 'Табачные изделия' },
                     { name: 'beer', label: 'Пиво' },
                     { name: 'alcohol', label: 'Алкогольная продукция' }
                  ]}
                  rules={{ required: 'Обязательное поле!' }}
                  required
               />
            )}

            {hasTypeField && (
               <RHFSelect
                  size="small"
                  name="type"
                  label="Тип ввода"
                  control={control}
                  options={[
                     { name: 'withQuotes', label: 'С кавычками' },
                     { name: 'withoutQuotes', label: 'Без кавычек' }
                  ]}
                  rules={{ required: 'Обязательное поле!' }}
                  required
               />
            )}

            <Box className="flex gap-2">
               <IconButton type="submit" color="primary">
                  <Search />
               </IconButton>
               <IconButton
                  onClick={() => {
                     onReset()
                     reset()
                  }}
                  color="secondary">
                  <Replay />
               </IconButton>
            </Box>
         </Box>
      </Box>
   )
}

export default Form
