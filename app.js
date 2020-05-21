const fs = require('fs')
const xlsx = require('xlsx')

fs.readFile('input.txt', 'utf8', (err, data) => {
    const text = data
    const words = text.split(' ')
    let pairs = []

    for (let x = 0; x < words.length; x++) {
        let word = words[x];
        word = word.replace(/\r?\n|\r/g, '')
        word = word.replace(/[.,\/#!$%\^&\*;:?{}=\-_`~()]/g, '')
        if (word === '' || word === '-') continue
        word = word.toLowerCase()
        
        const index = pairs.findIndex(el => el[0] === word)
        if (index !== -1) {
            pairs[index][1] += 1
        } else {
            pairs.push([word, 1])
        }
    }

    pairs = pairs.filter(el => el[1] > 1)

    pairs.sort((a, b) => {
        return b[1] - a[1]
    })
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(pairs)
    xlsx.utils.book_append_sheet(wb, ws, 'pairs');

    xlsx.writeFile(wb, 'out.xlsb')

})
