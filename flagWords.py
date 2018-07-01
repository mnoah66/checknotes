

def flaggedWords(self, ws, my_list):
        '''Finds keywords in row of data, throws in list'''
        for row in ws.iter_rows(row_offset=1):
            d = row[0] # The note
            e = row[1] # The name of the individual
            f = row[2] # Contact date
            g = row[3] # Program
            h = row[4] # Start time
            i = row[5] # end time
            j = row[6] # duration
            k = row[7] # Note writer
            foundWords = []
            if d.value:
                for w in sorted(my_list):
                    if w.lower() in str(d.value).lower():
                        foundWords.append(w)
                if len(foundWords) > 0:
                    note = ''
                    for l in foundWords:
                        left,sep,right = d.value.lower().partition(l)
                        note = note + left[-70:] + sep.upper() + right[:70] + ';'
                    forCSV = ','.join(foundWords).upper()
                    self.csvWritee(forCSV, e, h, i, f, note, g, j, k)