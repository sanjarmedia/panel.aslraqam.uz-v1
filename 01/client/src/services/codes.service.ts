import axios from 'axios'
import { ICodesHistory } from '../types/codesHistory.types'
import { ICodesInfo } from '../types/codesInfo.types'

class CodesService {
   private CODES_URL = ' http://localhost:8080/api'

   async getCodesHistory(id: string) {
      const encodedId = encodeURIComponent(id)
      const { data } = await axios.get<ICodesHistory>(`${this.CODES_URL}/history?id=${encodedId}`)
      return data
   }

   async getCodesInfo(codes: string[], pg: string) {
      const { data } = await axios.post<ICodesInfo[]>(`${this.CODES_URL}/info`, { codes, pg })
      return data
   }
}

export const codesService = new CodesService()
