require 'httparty'
require 'axlsx'
require 'json'

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
      sheet.add_row ["ID", "Début", "Fin", "Commentaire", "Créé à", "Mis à jour à"]

      sessions.each do |s|
        sheet.add_row [
                        s["id"],
                        s["started_at"],
                        s["ended_at"],
                        s["commentaire"],
                        s["created_at"],
                        s["updated_at"]
                      ]
      end
    end
    p.serialize("sessions.xlsx")
  end

  puts "✅ Fichier sessions.xlsx généré."
else
  puts "❌ Erreur HTTP #{response.code}"
end