import frappe
from frappe.utils.pdf import get_pdf
from frappe import _
import io
import requests
from docx import Document
from bs4 import BeautifulSoup
from docx.shared import Inches
from datetime import datetime


@frappe.whitelist()
def export_portfolio(portfolio_names, file_format="pdf"):
	if not portfolio_names:
		frappe.throw(_("No portfolio names provided"))

	portfolio_names = frappe.parse_json(portfolio_names)
	file_data_list = []

	for docname in portfolio_names:
		portfolio = frappe.get_doc("Portfolio", docname)
		content = generate_html_content(portfolio)

		if file_format == "pdf":
			file_data = get_pdf(content)
			file_extension = "pdf"
		elif file_format == "docx":
			file_data = generate_docx(content)
			file_extension = "docx"
		else:
			frappe.throw(_("Unsupported file format"))

		file_data_list.append(file_data)

	# Combine file_data_list into a single file if needed, for simplicity, return the first file.
	combined_file_data = file_data_list[0]

	# Generate a default file name with timestamp
	timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
	default_file_name = f"portfolio_export_{timestamp}.{file_extension}"

	# Create a new File document to save the generated file
	file_doc = frappe.get_doc({
		"doctype": "File",
		"file_name": default_file_name,
		"is_private": 1,
		"content": combined_file_data
	})
	file_doc.insert()

	return {
		"status": "success",
		"message": f"Portfolios exported successfully.",
		"file_url": file_doc.file_url
	}


def generate_html_content(portfolio):
	project_details = ""

	technologies_list = ""
	for tech in portfolio.technologies:
		technologies_list += f"<li>{tech.technology_name}</li>"

	images_list = ""
	for image in portfolio.images:
		if image:
			images_list += f'<img src="{image.website_image}" alt="Screenshot" style="width:300px;height:auto;"><br>'

	project_details += f"""
    <h3>{portfolio.title}</h3>
    {images_list}
    <p>Client: {portfolio.client}</p>
    <p>Period: {portfolio.start_date} - {portfolio.end_date}</p>
    <p>Technologies:</p>
    <ul>
    {technologies_list}
    </ul>
    <p>Details: {portfolio.body}</p>
    <p>URL: {portfolio.website}</p>
    <p>Location: {portfolio.location}</p>
    <hr>
    """

	content = f"""
    <html>
    <head>
        <title>Portfolio: PAST AND CURRENT PROJECTS</title>
    </head>
    <body>
        <h1>PAST AND CURRENT PROJECTS</h1>
        <p>Here is a selection of projects we have worked on or are working on.</p>
        {project_details}
    </body>
    </html>
    """
	return content


def generate_docx(html_content):
	doc = Document()
	soup = BeautifulSoup(html_content, 'html.parser')

	for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'li', 'img']):
		if element.name.startswith('h'):
			doc.add_heading(element.get_text(), level=int(element.name[1]))
		elif element.name == 'p':
			doc.add_paragraph(element.get_text())
		elif element.name == 'ul':
			for li in element.find_all('li'):
				doc.add_paragraph(f'• {li.get_text()}', style='ListBullet')
		elif element.name == 'img':
			response = requests.get(element['src'])
			img = io.BytesIO(response.content)
			doc.add_picture(img, width=Inches(4.0))

	output = io.BytesIO()
	doc.save(output)
	return output.getvalue()
