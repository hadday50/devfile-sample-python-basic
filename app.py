from flask import Flask, send_file, render_template_string, request
import os
import requests
from io import BytesIO
from openpyxl import Workbook

app = Flask(__name__)

MMD_URL = "https://esb.goteborg.se/MMD005_MasterData/Organisations?organisationType=FS"


@app.route("/")
def home():
    return """
    <h1>MMD - Förskolor</h1>
    <form method="get" action="/view">
    <input type="text" name="q" placeholder="Sök namn eller skolenhetskod">
    <button type="submit">Sök</button>
</form>
    <p>Hämtar förskolor från MMD och skapar en Excel-fil.</p>
    <a href="/download">Hämta Excel-fil</a>
    <br>
    <a href="/view">Visa förskolor i webbläsaren</a>
    """
@app.route("/view")
def view_data():
    username = os.environ.get("MMD_USERNAME")
    password = os.environ.get("MMD_PASSWORD")

    response = requests.get(
        MMD_URL,
        auth=(username, password),
        timeout=60
    )

    response.raise_for_status()
    json_data = response.json()
    query = request.args.get("q", "").strip().lower()
    organisations = json_data["data"]
    filtered_organisations = []
    for item in organisations:
        organisation = item.get("organisation") or {}
        name = organisation.get("displayName", "").lower()
        identifiers = organisation.get("organisationIdentifiers") or []
        skolenhetskod = ""

        for identifier in identifiers:
            if identifier.get("type") == "Skolenhetskod":
                skolenhetskod = str(identifier.get("value", "")).lower()

        if not query or query in name or query in skolenhetskod:
            filtered_organisations.append(item)

  return render_template_string("""
<h1>MMD - Förskolor</h1>

<p>Antal förskolor: {{ organisations|length }}</p>

<p>
    <a href="/">Till startsidan</a> |
    <a href="/download">Hämta Excel-fil</a>
</p>

<table border="1" cellpadding="6">
    <tr>
        <th>Namn</th>
        <th>Skolenhetskod</th>
        <th>HSA</th>
        <th>AltID</th>
        <th>Organisation ID</th>
        <th>OrganisationCode</th>
        <th>Startdatum</th>
        <th>Slutdatum</th>
    </tr>

    {% for item in organisations %}
    {% set organisation = item.get('organisation') or {} %}
    {% set identifiers = organisation.get('organisationIdentifiers') or [] %}
    {% set ns = namespace(skolenhetskod='', hsa='', altid='') %}
    {% for identifier in identifiers %}
    {% if identifier.get('type') == 'Skolenhetskod' %}
    {% set ns.skolenhetskod = identifier.get('value', '') %}
    {% elif identifier.get('type') == 'HSA' %}
        {% set ns.hsa = identifier.get('value', '') %}
    {% elif identifier.get('type') == 'AltID' %}
        {% set ns.altid = identifier.get('value', '') %}
        {% endif %}
        {% endfor %}
    <tr>
        <td>{{ organisation.get('displayName', '') }}</td>
        <td>{{ ns.skolenhetskod }}</td>
        <td>{{ ns.hsa }}</td>
        <td>{{ ns.altid }}</td>
        <td>{{ organisation.get('id', '') }}</td>
        <td>{{ organisation.get('organisationCode', '') }}</td>
        <td>{{ organisation.get('startDate', '') }}</td>
        <td>{{ organisation.get('endDate', '') }}</td>
    </tr>
    {% endfor %}
</table>
""", organisations=filtered_organisations)

@app.route("/download")
def download_excel():

    username = os.environ.get("MMD_USERNAME")
    password = os.environ.get("MMD_PASSWORD")

    if not username or not password:
        return "MMD_USERNAME eller MMD_PASSWORD saknas i OpenShift.", 500

    response = requests.get(
        MMD_URL,
        auth=(username, password),
        timeout=60
    )

    response.raise_for_status()

    json_data = response.json()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Förskolor"

    sheet.append([
        "Namn",
        "Skolenhetskod",
        "HSA",
        "AltID",
        "Organisation ID",
        "OrganisationCode",
        "Startdatum",
        "Slutdatum"
    ])

    for item in json_data.get("data", []):
        organisation = item.get("organisation", {})

        identifiers = {}

        for identifier in organisation.get("organisationIdentifiers", []):
            identifiers[identifier.get("type")] = identifier.get("value")

        sheet.append([
            organisation.get("displayName"),
            identifiers.get("Skolenhetskod"),
            identifiers.get("HSA"),
            identifiers.get("AltID"),
            organisation.get("id"),
            organisation.get("organisationCode"),
            organisation.get("startDate"),
            organisation.get("endDate")
        ])

    excel_file = BytesIO()
    workbook.save(excel_file)
    excel_file.seek(0)

    return send_file(
        excel_file,
        as_attachment=True,
        download_name="MMD_forskolor.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    port = int(os.environ.get("FLASK_PORT") or 8080)
    app.run(port=port, host="0.0.0.0")
