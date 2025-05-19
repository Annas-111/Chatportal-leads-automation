import 'dotenv/config'
import fs from 'fs'
import path from 'path'
import { google } from 'googleapis'
import ExcelJS from 'exceljs'
import { createClient } from '@supabase/supabase-js'

// Initialize Supabase
const supabase = createClient(
  process.env.SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY!
)

// Google Drive service account auth
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH!,
  scopes: ['https://www.googleapis.com/auth/drive.file']
})
const drive = google.drive({ version: 'v3', auth })

const LAST_EXPORT_FILE = path.resolve(__dirname, 'last-export.json')

// Read last export times from file or initialize empty object
function readLastExportTimes(): Record<string, string> {
  try {
    const json = fs.readFileSync(LAST_EXPORT_FILE, 'utf8')
    return JSON.parse(json)
  } catch {
    return {}
  }
}

// Write last export times back to file
function writeLastExportTimes(data: Record<string, string>) {
  fs.writeFileSync(LAST_EXPORT_FILE, JSON.stringify(data, null, 2))
}

// Helper to get folder ID
async function getOrCreateFolder(folderName: string): Promise<string> {
  const res = await drive.files.list({
    q: `mimeType='application/vnd.google-apps.folder' and name='${folderName}' and trashed=false`,
    fields: 'files(id, name)',
    spaces: 'drive'
  })

  const folder = res.data.files?.[0]
  if (folder) return folder.id!

  const created = await drive.files.create({
    requestBody: {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder'
    },
    fields: 'id'
  })

  return created.data.id!
}

// Find file by name in folder
async function findFileInFolder(fileName: string, folderId: string): Promise<string | null> {
  const res = await drive.files.list({
    q: `'${folderId}' in parents and name='${fileName}' and trashed=false`,
    fields: 'files(id, name)'
  })

  return res.data.files?.[0]?.id ?? null
}

// Upload or update Excel file
async function uploadOrUpdateFile(filePath: string, fileName: string, folderId: string) {
  const existingFileId = await findFileInFolder(fileName, folderId)

  const media = {
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    body: fs.createReadStream(filePath)
  }

  if (existingFileId) {
    await drive.files.update({
      fileId: existingFileId,
      media
    })
    console.log(`♻️ Updated: ${fileName} in folder ID ${folderId}`)
    console.log('')
  } else {
    await drive.files.create({
      requestBody: {
        name: fileName,
        parents: [folderId]
      },
      media,
      fields: 'id, name'
    })
    console.log(`✅ Uploaded new file: ${fileName} to folder ID ${folderId}`)
    console.log('')
  }

  fs.unlinkSync(filePath)
}

// Export data from Supabase and write to Excel
async function exportTableToExcel(
  table: string,
  fileName: string,
  folderName: string,
  columnOrder?: string[]
) {
  const lastExportTimes = readLastExportTimes()
  const lastExportTime = lastExportTimes[table] || '1970-01-01T00:00:00Z'

  // Fetch only new or updated rows since last export time
  const { data, error } = await supabase
    .from(table)
    .select('*')
    .or(`created_at.gte.${lastExportTime},updated_at.gte.${lastExportTime}`)

  console.log(`Fetched ${data?.length || 0} rows from ${table} since ${lastExportTime}`)
  console.log('')
  if (error) {
    console.error(`Error fetching from ${table}:`, error.message)
    console.log('')
    return
  }

  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet(table)

  if (data?.length) {
    const headers = columnOrder || Object.keys(data[0])
    sheet.addRow(headers)
    data.forEach(row => {
      sheet.addRow(headers.map(h => row[h]))
    })
  } else {
    console.log(`No new or updated data for ${table}. Skipping file write.`)
    console.log('')
    return
  }

  const localFilePath = path.resolve(__dirname, fileName)
  await workbook.xlsx.writeFile(localFilePath)

  // Get folder ID from env based on folderName param
  let folderId = ''
  if (folderName === 'agentsForm') folderId = process.env.AGENTS_FOLDER_ID!
  else if (folderName === 'contactFormData') folderId = process.env.CONTACT_FORM_FOLDER_ID!
  else if (folderName === 'chatbotsForm') folderId = process.env.CHATBOTS_FOLDER_ID!
  else folderId = await getOrCreateFolder(folderName)

  await uploadOrUpdateFile(localFilePath, fileName, folderId)

  // Update last export time
  lastExportTimes[table] = new Date().toISOString()
  writeLastExportTimes(lastExportTimes)
}

// Run exports
async function run() {
  await exportTableToExcel(
    'agents',
    'agents-data.xlsx',
    'agentsForm',
    [
      'id',
      'user_id',
      'company_name',
      'company_type',
      'agent_type',
      'created_at',
      'updated_at'
    ]
  )

  await exportTableToExcel(
    'contact_us_submissions',
    'contact-submissions.xlsx',
    'contactFormData',
    ['id', 'name', 'email', 'purpose', 'created_at', 'updated_at']
  )

  await exportTableToExcel(
    'chatbots',
    'chatbots-data.xlsx',
    'chatbotsForm',
    [
      'id',
      'user_id',
      'company_name',
      'website',
      'page_limit',
      'company_type',
      'agent_type',
      'status',
      'created_at',
      'updated_at'
    ]
  )
}

run().catch(err => {
  console.error(err)
  console.log('')
  process.exit(1)
})
