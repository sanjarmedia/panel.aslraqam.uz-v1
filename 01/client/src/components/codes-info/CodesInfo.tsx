import { Box } from '@mui/material'
import { DataGrid, GridColDef, useGridApiRef } from '@mui/x-data-grid'
import { useQueries } from '@tanstack/react-query'
import delay from 'delay'
import { FC, useState } from 'react'
import { FormProvider, useForm } from 'react-hook-form'

import { codesService } from '../../services/codes.service'
import { chunkArray } from '../../utils/chunkArray'
import { validateCodes } from '../../utils/validateCodes'

import ExportButton from '../ui/ExportButton'
import Form from '../ui/Form'

const CodesInfo: FC = () => {
   const [info, setInfo] = useState<{ pg: string; codes: string[] }>({ pg: '', codes: [] })
   const methods = useForm<{ codeList: string; pg: string; type: string }>()
   const apiRef = useGridApiRef()

   const codeChunks = chunkArray(info.codes, 1000)
   const queries = useQueries({
      queries: codeChunks.map((chunk, idx) => ({
         queryKey: ['codesInfo', chunk, info.pg],
         queryFn: async () => {
            await delay(idx * 100)
            return codesService.getCodesInfo(chunk, info.pg)
         },
         enabled: chunk.length > 0
      }))
   })

   const onSubmit = (data: { codeList: string; pg?: string; type?: string }, limit: number) => {
      const { newCodes, error } = validateCodes(data.codeList, data.type === 'withQuotes', limit)
      if (error) return methods.setError('codeList', { message: error })

      setInfo((prev) => ({ ...prev, pg: data.pg!, codes: [...new Set([...prev.codes, ...newCodes!])] }))
      methods.clearErrors('codeList')
      methods.reset({ codeList: '', pg: data.pg, type: data.type })
   }

   const columns: GridColDef[] = [
      { field: 'requestedCis', headerName: 'CIS', sortable: false, flex: 1 },
      {
         field: 'child',
         headerName: 'Childs',
         flex: 1,
         valueGetter: (params: string[]) => (params ? params.length : '')
      },
      { field: 'parent', headerName: 'Parent', flex: 1 },
      {
         field: 'expirationDate',
         headerName: 'EX date',
         flex: 1,
         valueFormatter: (params: string) => (params ? new Date(params).toLocaleDateString('ru-RU') : '')
      },
      { field: 'ownerInn', headerName: 'TIN', flex: 1 },
      { field: 'status', headerName: 'Status', flex: 1 },
      { field: 'statusEx', headerName: 'AIC', flex: 1 }
   ]

   const rows = queries.flatMap((obj) => obj?.data ?? []).map((i, idx) => ({ id: idx + 1, ...i.cisInfo }))
   console.log(info.codes)

   return (
      <Box className="mt-8 space-y-5">
         <FormProvider {...methods}>
            <Form onSubmit={(data) => onSubmit(data, 50000)} onReset={() => setInfo({ pg: '', codes: [] })} hasPGField hasTypeField />
         </FormProvider>

         <Box className="overflow-x-auto">
            <Box className="h-[600px] min-w-[1000px]">
               <DataGrid
                  rows={rows}
                  columns={columns}
                  loading={queries.some((q) => q.isLoading)}
                  apiRef={apiRef}
                  initialState={{ pagination: { paginationModel: { page: 0, pageSize: 100 } } }}
                  pageSizeOptions={[100]}
                  checkboxSelection
                  disableRowSelectionOnClick
                  sx={{ '& .MuiDataGrid-columnHeaderTitle': { fontWeight: 'bold' } }}
               />
            </Box>
         </Box>

         <ExportButton apiRef={apiRef} type="application/json" fileName="info.json" sx={{ textTransform: 'capitalize' }}>
            Export JSON
         </ExportButton>
      </Box>
   )
}

export default CodesInfo
