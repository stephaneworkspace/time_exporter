require 'httparty'
require 'axlsx'
require 'json'
require 'time'

TOKEN = File.read("token.txt").strip

response = HTTParty.get(
  'https://time.bressani.dev:3443/api/sessions',
  headers: {
    'Authorization' => "Bearer #{TOKEN}",
    'Accept' => '*/*'
  }
)

if response.code == 200
  sessions = JSON.parse(response.body).select { |s| s["project_id"] == 5 }

  Axlsx::Package.new do |p|
    p.workbook.add_worksheet(name: '5') do |sheet|
      sheet.add_row ["ID", "Début", "Fin", "Commentaire"]

      styles = sheet.workbook.styles
      datetime_style = styles.add_style(format_code: "dd/mm/yyyy hh:mm:ss")

      sessions.each do |s|
        sheet.add_row [
                        s["id"],
                        Time.parse(s["started_at"]),
                        Time.parse(s["ended_at"]),
                        s["commentaire"]
                      ], style: [nil, datetime_style, datetime_style, nil]
      end
    end
    p.serialize("sessions.xlsx")
  end

  puts "✅ Fichier sessions.xlsx généré."
else
  puts "❌ Erreur HTTP #{response.code}"
end