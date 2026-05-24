export const chunkArray = (array: string[], size: number): string[][] => {
   const result: string[][] = []
   for (let i = 0; i < array.length; i += size) {
      result.push(array.slice(i, i + size))
   }
   return result
}
