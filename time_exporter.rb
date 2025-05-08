EXPORT_DIR = File.expand_path('~/kDrive/Time/Sessions')
Dir.mkdir(EXPORT_DIR) unless Dir.exist?(EXPORT_DIR)
require 'httparty'
require 'axlsx'
require 'json'
require 'time'
require 'tzinfo'

TOKEN = File.read("token.txt").strip

categories_response = HTTParty.get(
  'https://time.bressani.dev:3443/api/categories',
  headers: {
    'Authorization' => "Bearer #{TOKEN}",
    'Accept' => '*/*'
  }
)

def generate_excel_from_sessions(response, category_name)
  if response.code == 200
    sessions_by_project = JSON.parse(response.body).group_by { |s| s.dig("project", "name") }

    safe_name = category_name.gsub(/[\/\\:*?"<>|]/, '_')
    filename = File.join(EXPORT_DIR, "Sessions - #{safe_name}.xlsx")

    Axlsx::Package.new do |p|
      sessions_by_project.each do |project_name, sessions|
        p.workbook.add_worksheet(name: project_name[0..30]) do |sheet|
          styles = sheet.workbook.styles
          datetime_bordered_style = styles.add_style(
            format_code: "dd/mm/yyyy hh:mm:ss",
            border: { style: :thin, color: "000000", edges: [:top, :bottom, :left, :right] }
          )
          decimal_bordered_style = styles.add_style(
            format_code: "0.00",
            border: { style: :thin, color: "000000", edges: [:top, :bottom, :left, :right] }
          )
          bordered_style = styles.add_style(
            border: { style: :thin, color: "000000", edges: [:top, :bottom, :left, :right] }
          )
          header_style = styles.add_style(
            b: true,
            border: {
              style: :medium,
              color: "000000",
              edges: [:top, :bottom, :left, :right]
            },
            alignment: { horizontal: :center }
          )

          sheet.add_row ["ID", "Début", "Fin", "Durée", "Jour travail", "Commentaire"], style: [header_style] * 6

          sessions.each do |s|
            tz = TZInfo::Timezone.get('Europe/Zurich')
            started_at = tz.utc_to_local(Time.parse(s["started_at"]).utc)
            ended_at = tz.utc_to_local(Time.parse(s["ended_at"]).utc)
            duration_seconds = ended_at - started_at
            duration = Time.at(duration_seconds).utc.strftime("%H:%M:%S")
            sheet.add_row [
              s["id"],
              started_at,
              ended_at,
              duration,
              nil,
              s["commentaire"]
            ], style: [bordered_style, datetime_bordered_style, datetime_bordered_style, bordered_style, decimal_bordered_style, bordered_style]

            last_row_index = sheet.rows.size
            sheet.rows.last.cells[4].value = "=(C#{last_row_index}-B#{last_row_index})*24/8.5"
          end

          total_row_index = sheet.rows.size + 1
          sheet.add_row ["", "", "", "", nil, "Total jours"], style: [bordered_style]*6
          sheet.rows.last.cells[4].value = "=SUM(E2:E#{total_row_index - 1})"

          sheet.add_row ["", "", "", "", nil, "Total heures"], style: [bordered_style]*6
          sheet.rows.last.cells[4].value = "=E#{total_row_index}*8.5"
        end
      end
      p.serialize(filename)
    end

    puts "✅ Fichier #{filename} généré."
  else
    puts "❌ Erreur HTTP #{response.code}"
  end
end

if categories_response.code == 200
  JSON.parse(categories_response.body).each do |category|
    category_id = category["id"]
    category_name = category["name"]

    response = HTTParty.get(
      "https://time.bressani.dev:3443/api/sessions?category_id=#{category_id}",
      headers: {
        'Authorization' => "Bearer #{TOKEN}",
        'Accept' => '*/*'
      }
    )

    generate_excel_from_sessions(response, category_name)
  end
else
  puts "❌ Erreur lors de la récupération des catégories (HTTP #{categories_response.code})"
end