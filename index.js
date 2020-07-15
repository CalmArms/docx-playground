const mammoth = require('mammoth')
const fs = require('fs')
const html2json = require('html2json').html2json

async function convertDocXFileToHTML(filePath, fileName) {
    const result = await mammoth.convertToHtml({ path: filePath }).catch(e => e)
    if (result instanceof Error) {
        console.warn(`[error] unable to convert ${fileName}.docx into HTML`)
        throw result
    } else {
        result.messages.forEach(message => console.warn(message))
        fs.writeFile(`./output/${fileName}.html`, result.value, e => {
            if (e) throw e
        })
        return result.value
    }
}

function convertHTMLToJSON(html, fileName) {
    const json = html2json(html)
    const jsonString = JSON.stringify(json)
    fs.writeFile(`./output/${fileName}.json`, jsonString, e => {
        if (e) console.error(e)
    })
    return json
}

async function convertDocXFileToJSON(filePath, fileName) {
    const html = await convertDocXFileToHTML(filePath, fileName)
    if (html) {
        convertHTMLToJSON(html, fileName)
    }
}

console.log('[info] starting to convert files...')

fs.readdir('./files', { withFileTypes: true }, async function(err, files) {
    console.log(`[info] found ${files.length} file(s)`)

    let hiddenFilesCount = 0
    let nonDocXFilesCount = 0
    let failedConversions = []

    if (err) {
        console.error(err)
    } else {
        files = files.filter(file => {
            // filter out hidden files
            if ((/(^|\/)\.[^\/\.]/g).test(file.name)) {
                hiddenFilesCount++
                return false
            }
            // filter out non .docx files
            if (!file.name.includes('.docx')) {
                nonDocXFilesCount++
                return false
            }
            return true
        })

        if (hiddenFilesCount > 0) {
            console.log(`[info] not converting ${hiddenFilesCount} hidden file(s)`)
        }

        if (nonDocXFilesCount > 0) {
            console.log(`[info] not converting ${nonDocXFilesCount} non .docx file(s)`)
        }
       
        console.log(`[info] attempting to convert ${files.length} .docx file(s)`)

        await Promise.all(files.map(async file => {
            const fileName = file.name.split('.docx')[0]
            const filePath = `./files/${file.name}`
            try {
                await convertDocXFileToJSON(filePath, fileName)
            } catch (e) {
                failedConversions.push(fileName)
                // console.error(e)
            }
        }))

        console.log(`[info] completed converting files.`)
        console.log(`[info] successfully converted ${files.length - failedConversions.length} files`)
        console.log(`[info] failed to convert the following files:`, failedConversions)
    }
})