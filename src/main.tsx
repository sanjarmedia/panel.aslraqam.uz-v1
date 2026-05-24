import { createRoot } from 'react-dom/client'
import { Provider } from './utils/Provider.tsx'

import App from './App.tsx'
import './main.css'

createRoot(document.getElementById('root')!).render(
   <Provider>
      <App />
   </Provider>
)
