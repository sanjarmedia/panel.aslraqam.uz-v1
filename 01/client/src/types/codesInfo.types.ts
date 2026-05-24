export interface ICodesInfo {
   cisInfo: CisInfoSuccess | CisInfoError
}

interface CisInfoSuccess {
   requestedCis: string
   cis: string
   parent?: string
   child?: string[]
   status: string
   packageType: string
   productGroupId: number
   productGroup: string
   gtin: string
   ownerInn: string
   ownerName: string
   statusEx: string
   productionSerialNumber: string
   expirationDate: string
}

interface CisInfoError {
   requestedCis: string
   errorMessage: string
   errorCode: string
}
