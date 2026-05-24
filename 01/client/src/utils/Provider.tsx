import { QueryClient, QueryClientProvider } from '@tanstack/react-query'
import { type PropsWithChildren, useState } from 'react'

export function Provider({ children }: PropsWithChildren) {
  const [client] = useState(
    new QueryClient({
      defaultOptions: {
        queries: { refetchOnWindowFocus: false }
      }
    })
  )

  return <QueryClientProvider client={client}>{children}</QueryClientProvider>
}
