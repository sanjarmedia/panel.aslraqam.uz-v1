import { CodesHistorySuccess } from '../types/codesHistory.types'

export const valueFormatter = (data: CodesHistorySuccess[], field: keyof CodesHistorySuccess) =>
   data
      .map((item) => {
         if (field === 'timestamp' && item[field]) return new Date(item[field]).toLocaleDateString('ru-RU')
         if (field === 'child' && item[field]) return item[field].length
         return item[field] || '-'
      })
      .join('\n')
