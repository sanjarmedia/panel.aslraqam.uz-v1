import { Box } from '@mui/material'
import { DataGrid, GridColDef, useGridApiRef } from '@mui/x-data-grid'
import { useQueries } from '@tanstack/react-query'
import delay from 'delay'
import { FC, useState } from 'react'
import { FormProvider, useForm } from 'react-hook-form'

import { codesService } from '../../services/codes.service'
import { validateCodes } from '../../utils/validateCodes'

import { valueFormatter } from '../../utils/valueFormatter'
import ExportButton from '../ui/ExportButton'
import Form from '../ui/Form'

const CodesHistory: FC = () => {
   const [history, setHistory] = useState<string[]>([])
   const methods = useForm<{ codeList: string; type: string }>()
   const apiRef = useGridApiRef()

   const queries = useQueries({
      queries: history.map((id, idx) => ({
         queryKey: ['codesHistory', id],
         queryFn: async () => {
            await delay(idx * 100)
            return codesService.getCodesHistory(id)
         },
         enabled: history.length > 0
      }))
   })

   const onSubmit = (data: { codeList: string; pg?: string; type?: string }, limit: number) => {
      const { newCodes, error } = validateCodes(data.codeList, data.type === 'withQuotes', limit)
      if (error) return methods.setError('codeList', { message: error })

      setHistory((prev) => [...new Set([...prev, ...newCodes!])])
      methods.clearErrors('codeList')
      methods.reset({ codeList: '', type: data.type })
   }

   const columns: GridColDef[] = [
      { field: 'cis', headerName: 'CIS', flex: 1, sortable: false },
      { field: 'child', headerName: 'Childs', flex: 1 },
      { field: 'parent', headerName: 'Parent', flex: 1 },
      { field: 'timestamp', headerName: 'Date', flex: 1 },
      { field: 'ownerInn', headerName: 'TIN', flex: 1 },
      { field: 'status', headerName: 'Status', flex: 1 },
      { field: 'docId', headerName: 'Doc ID', width: 300 }
   ]

   const rows = queries.map((obj, idx) => {
      const data = obj.data

      if (Array.isArray(data)) {
         return {
            id: idx + 1,
            cis: data[0].cis,
            parent: valueFormatter(data, 'parent'),
            child: valueFormatter(data, 'child'),
            ownerInn: valueFormatter(data, 'ownerInn'),
            timestamp: valueFormatter(data, 'timestamp'),
            status: valueFormatter(data, 'status'),
            docId: valueFormatter(data, 'docId')
         }
      }

      return {
         id: idx + 1,
         cis: data?.cis,
         parent: '',
         child: '',
         ownerInn: '',
         timestamp: '',
         status: '',
         docId: ''
      }
   })

   return (
      <Box className="mt-8 space-y-5">
         <FormProvider {...methods}>
            <Form onSubmit={(data) => onSubmit(data, 1000)} onReset={() => setHistory([])} hasTypeField />
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
                  getRowHeight={(params) => params.model.docId.length || 30}
                  checkboxSelection
                  disableRowSelectionOnClick
                  virtualizeColumnsWithAutoRowHeight
                  clipboardCopyCellDelimiter={'\n'}
                  sx={{
                     '& .MuiDataGrid-columnHeaderTitle': { fontWeight: 'bold' },
                     '& .MuiDataGrid-cell': { whiteSpace: 'pre-wrap', lineHeight: 2 }
                  }}
               />
            </Box>
         </Box>

         <ExportButton apiRef={apiRef} type="application/json" fileName="history.json" sx={{ textTransform: 'capitalize' }}>
            Export JSON
         </ExportButton>
      </Box>
   )
}

export default CodesHistory
