const express = require('express')
const Anthropic = require('@anthropic-ai/sdk')
const XLSX = require('xlsx')

const app = express()
app.use(express.json({ limit: '20mb' }))

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

// Health check
app.get('/', (req, res) => res.json({ status: 'ok', service: 'arthvo-parser' }))

const STATEMENT_TOOL = {
  name: 'submit_bank_statement',
  description: 'Submit parsed bank statement data',
  input_schema: {
    type: 'object',
    properties: {
      bank: { type: 'string' },
      accountHolder: { type: 'string' },
      period: { type: 'string' },
      openingBalance: { type: 'number' },
      closingBalance: { type: 'number' },
      totalCredits: { type: 'number' },
      totalDebits: { type: 'number' },
      transactions: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            date: { type: 'string' },
            description: { type: 'string' },
            amount: { type: 'number' },
            type: { type: 'string', enum: ['credit', 'debit'] },
            category: { type: 'string', enum: ['salary','rent','emi','grocery','food','fuel','shopping','entertainment','insurance','investment','sip','transfer','utility','medical','education','other'] }
          },
          required: ['date','description','amount','type','category']
        }
      },
      summary: {
        type: 'object',
        properties: {
          salary: { type: 'number' }, rent: { type: 'number' }, emi: { type: 'number' },
          grocery: { type: 'number' }, food: { type: 'number' }, fuel: { type: 'number' },
          shopping: { type: 'number' }, entertainment: { type: 'number' }, insurance: { type: 'number' },
          investment: { type: 'number' }, sip: { type: 'number' }, utility: { type: 'number' },
          medical: { type: 'number' }, education: { type: 'number' }, other: { type: 'number' }
        },
        required: ['salary','rent','emi','grocery','food','fuel','shopping','entertainment','insurance','investment','sip','utility','medical','education','other']
      }
    },
    required: ['bank','accountHolder','period','openingBalance','closingBalance','totalCredits','totalDebits','transactions','summary']
  }
}

const SYSTEM = `You are a precise Indian bank statement parser. Extract ALL transactions from the bank statement provided.
Use the submit_bank_statement tool to return the parsed data. Categorise every transaction:
- salary: SALARY, NEFT from employer, payroll
- rent: RENT, house rent  
- emi: EMI, loan repayment, NACH
- sip: SIP, mutual fund, ZERODHA, GROWW, KUVERA
- investment: RD, FD, PPF, NPS contributions
- food: Swiggy, Zomato, restaurants, hotels, cafes
- grocery: BigBasket, DMart, supermarkets
- fuel: petrol pump, HPCL, BPCL, Indian Oil, Ola/Uber/Rapido
- entertainment: Netflix, Hotstar, Spotify, movies, gaming
- shopping: Amazon, Flipkart, Myntra, retail
- utility: electricity, gas, internet, mobile
- medical: pharmacy, hospital, doctor
- education: school fees, courses
- transfer: UPI to individuals, NEFT/IMPS to persons
- other: everything else
All amounts as plain numbers. Sum per category for summary.`

async function decryptExcel(buffer, password) {
  // Try without password first
  try {
    return XLSX.read(buffer, { type: 'buffer', password: password || undefined })
  } catch (e) {
    if (!password) throw new Error('requires_password')
    // Try officecrypto-tool for encrypted Excel
    try {
      const officeCrypto = require('officecrypto-tool')
      const decrypted = await officeCrypto.decrypt(buffer, { password })
      return XLSX.read(decrypted, { type: 'buffer' })
    } catch (e2) {
      throw new Error('incorrect_password')
    }
  }
}

function bufferToLines(buffer, password, fileKind) {
  // CSV
  if (fileKind === 'csv') {
    return buffer.toString('utf8')
  }
  // Image — return base64
  if (fileKind === 'image') {
    return null // handled separately
  }
  return null
}

app.post('/parse', async (req, res) => {
  const t0 = Date.now()
  const log = (msg) => console.log(`[${Date.now()-t0}ms] ${msg}`)

  try {
    const { base64, fileName, mimeType, password, fileKind } = req.body

    log(`Received: fileName=${fileName}, kind=${fileKind}, base64Length=${base64?.length}, hasPassword=${!!password}`)

    if (!base64) return res.status(400).json({ error: 'No file data' })
    if (!fileKind) return res.status(400).json({ error: 'No fileKind' })

    const buffer = Buffer.from(base64, 'base64')
    let claudeContent = []

    if (fileKind === 'excel-xlsx' || fileKind === 'excel-xls') {
      log('Parsing Excel...')
      let workbook
      try {
        workbook = await decryptExcel(buffer, password)
      } catch (e) {
        if (e.message === 'requires_password') return res.status(422).json({ error: 'incorrect_password' })
        if (e.message === 'incorrect_password') return res.status(422).json({ error: 'incorrect_password' })
        throw e
      }
      log('Excel decrypted')

      let csvText = ''
      for (const name of workbook.SheetNames) {
        const sheet = workbook.Sheets[name]
        const csv = XLSX.utils.sheet_to_csv(sheet)
        // Strip empty lines, keep only rows with numbers
        const lines = csv.split('\n').filter(l => {
          const clean = l.replace(/,+$/,'').trim()
          return clean && /\d/.test(clean)
        })
        if (lines.length) csvText += `\n=== ${name} ===\n${lines.join('\n')}\n`
      }
      log(`CSV ready: ${csvText.length} chars`)
      claudeContent = [{ type: 'text', text: `Parse this Indian bank statement:\n\n${csvText}` }]

    } else if (fileKind === 'pdf') {
      claudeContent = [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } },
        { type: 'text', text: 'Parse this Indian bank statement and extract all transactions.' }
      ]
    } else if (fileKind === 'image') {
      const mediaType = buffer[0] === 0xFF ? 'image/jpeg' : 'image/png'
      claudeContent = [
        { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64 } },
        { type: 'text', text: 'This is a bank statement photo. Extract all visible transactions.' }
      ]
    } else if (fileKind === 'csv') {
      const text = buffer.toString('utf8')
      claudeContent = [{ type: 'text', text: `Parse this bank statement CSV:\n\n${text}` }]
    } else {
      return res.status(415).json({ error: 'unsupported_format' })
    }

    log('Calling Claude...')
    const response = await client.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 16000,
      system: SYSTEM,
      tools: [STATEMENT_TOOL],
      tool_choice: { type: 'tool', name: 'submit_bank_statement' },
      messages: [{ role: 'user', content: claudeContent }]
    })
    log(`Claude done in ${Date.now()-t0}ms`)

    const toolUse = response.content.find(c => c.type === 'tool_use')
    if (!toolUse) return res.status(500).json({ error: 'No tool response from Claude' })

    return res.json({ data: toolUse.input, fileKind })
  } catch (err) {
    log(`Error: ${err.message}`)
    console.error(err)
    return res.status(500).json({ error: err.message || 'Failed to parse' })
  }
})

const PORT = process.env.PORT || 3001
app.listen(PORT, () => console.log(`ArthVo parser worker running on port ${PORT}`))
