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

// const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH!
const creds = {
  type: "service_account",
  project_id: "supabase-test-458622",
  private_key_id: "29941f7d6f3e45fe30b47b8d21c0f263a187e36d",
  private_key: `-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDvy52VJ5VueRQV\nkQkIYvUXfs3RvGke9+r1cJVC67RuNT3m5DCPFT3/g/JtPnfihEpQb5ahC+MS/qr3\nww2hEzFvw6OSMWvSE/sRgp/WqJUJIg5gQ4CoLlsdQmdG+xzRs9iya+fP6PZYuGWC\nczCIYiq94sWIzv8ItS2zGvQuUHdrW0HIl0C9LbOtRGbEg8pciv7LA1mzoC/gjlzs\nm/NTSMWwentaQUFMoZlYhxxzI1BT/n/qDKIQ7YfHyLDQGARuvUIGQHTuvGGPbEs3\nvXtyjJ060dx1n/Vu8K4v3FNpPFiQOFBkKybTXUXsi9WXMv7zagVE6PZxTx2ZrMBt\nzF5ksViPAgMBAAECggEACBLMau2qv83uya0Y50Hhq5lW16VmyqahMANK2zZxXDaJ\nr4DmS6L08+ne2yO7yJIYqG2hXim2yvhtDCtyhYZOZ0toce8dCLUoglXqCMGjRuvx\nHPeY2NC6u5j1PjQNK9hIxvUQGHVo+eK3dDVpoGXht4Pvj6QcfRAQilOxfe+ynZgj\nxteom+c18N7HHBAuWDTORlDuhbUPfBHhKgDgqC7BbOB4GS32igsuilia+Ux86UlY\nHReG9YDWngnmlqAsM1wZX3cQ83pzsymllnlE2DH9cE+zyBnJ8Kl99o/6JibsfSzl\nrtGXScrlGcPSKRS5gevTADnY1gMuB+/KCc9QcTNf8QKBgQD7dnwyb33M81uCwzLt\n+ZE3YL5AvyHRLb5BGbFspQIlvb3WFdXqMbElqspbQSv+kAMiiNlnr/Mo7AV+r7/5\nBFPhSm22AVNRbE0NQhHhmvNOlLnnCSzCDU4PNDhaAXgLx6floZHGIHsgZRXNgrNV\npOTSFLayDItV6oXocuXhXGyG3wKBgQD0HzzwiXVx4kJ2CSvrQBgQB6+ubom3urfV\nS/qiZxtz2EvgTfvgeR9fXyyeodCLR7GlJWABlUTXiM4z+sfoqeX1DxNDXCned4uE\nt+flFfqw8QbVtvw3cqPh+ouV2C67kmVGEpBQ0UJ0z5O91o6PTRbuYdAbOPcOYqgm\n3N44wJbUUQKBgQDS5/qc9qPlnQrYrs1tiU9ByjNn7Kb5YctboKgN/ovBidXJ1ICZ\nL1prxEZ6qTu0A6eNdIgbjhh87xBgxBlYS0frAljoOU5fUV2CzDJhLWv6MYWGgEx4\n6V3aJKtK2kaMYsWwNIYmTaHXdtEFkCTHpUiB5vugdCL9SXlMj+m8zZ78swKBgDhA\nHm2abtpv9Tp3gIwzd5fx/XQm+krAlm1qlXToZzX5R/qGXZxqrGTwbDqdNI1zVqak\nBww2VeeIGWN2zKt6wocGEz05NliRmS9apO9vchVlZFrIUDOCkeqXLaS4bIRCBl3w\nFxrYLQT3c6aRksXIUJlbJnWvWZYR+4cA65+OGPMRAoGBAN95SPVrxlDYFZ05MSBU\ndhHSVZWZbzKiqJuso4GzPbQVwQyZQ2hWuIhltkg3r+5H0FwhKALjPeYMPEuP9q6L\n9Mo3Bt7qiegV3PnX5keXeAMdkdN9ocM54C0PwWM/8zpHpq4I42imxtEoHy/7GjMd\nVpZg/Yw4fhaL9NYX8ujKRGrQ\n-----END PRIVATE KEY-----\n`,
  client_email: "gdrive-exporter@supabase-test-458622.iam.gserviceaccount.com",
  client_id: "101355721592912771983",
  auth_uri: "https://accounts.google.com/o/oauth2/auth",
  token_uri: "https://oauth2.googleapis.com/token",
  auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
  client_x509_cert_url:
    "https://www.googleapis.com/robot/v1/metadata/x509/gdrive-exporter%40supabase-test-458622.iam.gserviceaccount.com",
  universe_domain: "googleapis.com"
};

const auth = new google.auth.GoogleAuth({
  credentials: creds,
  scopes: ['https://www.googleapis.com/auth/drive.file'],
});
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
