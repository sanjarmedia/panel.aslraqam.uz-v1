import { Button, useMediaQuery } from '@mui/material'
import { GridApiCommunity } from '@mui/x-data-grid/internals'
import { saveAs } from 'file-saver'
import { FC } from 'react'

interface IExportButton {
   children: React.ReactNode
   apiRef: React.RefObject<GridApiCommunity>
   sx?: object
   type: string
   fileName: string
}

const ExportButton: FC<IExportButton> = ({ children, apiRef, sx, type, fileName }) => {
   const isMobile = useMediaQuery('(max-width: 450px)')

   const downloadAsJson = () => {
      const rows = Array.from(apiRef.current?.getSelectedRows().values() || [])
      saveAs(new Blob([JSON.stringify(rows, null, 2)], { type }), fileName)
   }

   return (
      <Button fullWidth={isMobile} sx={sx} onClick={downloadAsJson} variant="contained" disableElevation>
         {children}
      </Button>
   )
}

export default ExportButton
