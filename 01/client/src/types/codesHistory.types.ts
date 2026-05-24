export type ICodesHistory = CodesHistorySuccess[] | CodesHistoryError
export interface CodesHistorySuccess {
   cis: string
   gtin: string
   packageType: string
   ownerInn: string
   status: string
   parent?: string
   child?: string[]
   producerInn: string
   timestamp: string
   operationDate: string
   emissionDate: string
   docId: string
   productionDate?: string
   expirationDate?: string
   productionSerialNumber?: string
}

interface CodesHistoryError {
   cis: string
   error: string
}
