import { Code, History } from '@mui/icons-material'
import { Alert, Box, Tab, Tabs } from '@mui/material'
import { FC, useState } from 'react'

import CodesHistory from './components/codes-history/CodesHistory'
import CodesInfo from './components/codes-info/CodesInfo'

const tabs = [
   { name: 'codes-info', label: 'Info', icon: <Code /> },
   { name: 'codes-history', label: 'History', icon: <History /> }
]

const App: FC = () => {
   const [tab, setTab] = useState<number>(0)

   return (
      <Box className="mx-auto h-full max-w-[2000px] p-3">
         <Tabs value={tab} onChange={(_, newValue) => setTab(newValue)}>
            {tabs.map((obj) => (
               <Tab key={obj.name} label={obj.label} icon={obj.icon} iconPosition="start" sx={{ textTransform: 'none' }} />
            ))}
         </Tabs>

         <Alert severity="info">Список атрибутов необходимо указывать либо в кавычках через запятую, либо без кавычек, разделяя пробелами.</Alert>

         {tab === 0 && <CodesInfo />}
         {tab === 1 && <CodesHistory />}
      </Box>
   )
}

export default App
