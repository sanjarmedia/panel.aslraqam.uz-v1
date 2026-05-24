const cleanString = (str: string) => str.replace(/\\"/g, '"')

export const validateCodes = (data: string, withQuotes: boolean, limit: number) => {
   let newCodes: string[]

   if (withQuotes) {
      const matches = Array.from(data.matchAll(/(["'])(.*?[^\\])\1/g), (m) => cleanString(m[2]))
      if (!matches.length) return { error: 'Атрибуты должны быть в кавычках!' }
      newCodes = matches
   } else {
      if (/^["']/.test(data)) return { error: 'Атрибуты не должны начинаться с кавычек!' }

      newCodes = data
         .split(/\s+/)
         .map((c) => cleanString(c.trim()))
         .filter(Boolean)
   }

   if (newCodes.length > limit) {
      return { error: `Превышен лимит: максимум ${limit}, сейчас — ${newCodes.length}` }
   }

   return { newCodes }
}
