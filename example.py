from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt 
import datetime

# dynamic info
name = "Duc Nguyen"
address = "Victoria BC"
phone = "250-884-6325"
email = "dukeng@uvic.ca"
website="https://dukeng.github.io/"
github="https://github.com/dukeng/"

company_name="Relic Entertainment"
position="Associate Programmer (Server) Co-op"
open_para="I am writing to apply to the {} position with {}. I am a third year as a Computer Science student at the University of Victoria and I have very strong interests in software development; I can contribute to the position with my self-driven software knowledge as evidenced by my coop experience and my pursuit of various extracurricular activities. "
first_coop="In the summer of 2017, I worked as a software development engineer for Abebooks, an  Amazon Company for four months. I used Python to prototype the integration of a potential online payment service into the current systems to determine the feasibility for future business between the two companies. I tested the payment provider’s API, participated in evaluating risks and security loopholes and negotiating for customized features. Besides, I developed ecommerce features using Groovy with Spring MVC, JUnit, AWS: S3, SQS and databases such as PostgreSQL and Oracle. "
second_coop="My other work experience includes a four month coop term in September 2016 as a full stack web developer for RealtyServer System Inc., a company that offers multiple listing services for real estate business. My duties included implementing new features and functionalities for various real estate websites, using technology such as Java Servlet, Velocity template engine and Javascript, jQuery for frontend styling and RESTful API."
activities = [
    "I competed in 2018 Battlesnake AI competition to build a C++ snake to fight other snakes and won the first prize of $1000",
    "I created an Android game application which is published on GooglePlay and has achieved 500 downloads",
    "I used Chrome Extension API to create an extension that helps UVic students register for courses easily",
    "I have attended a number of hackathons and have created a web application using GoogleMap API, Python and PostgreSQL that won a Sponsor Prize at the UVic Hacks 2017"
    ]
closing = "I would like to thank you very much for considering my application, and I would greatly look forward to an opportunity to interview for the position offered by {}. I can be reached by phone at {} or by email at {}. For more information about me, please visit: {}" 

document = Document()

document.add_heading(name, 2)

margin = 2
#changing the page margins
sections = document.sections
for section in sections:
    section.top_margin = Cm(margin)
    section.bottom_margin = Cm(margin)
    section.left_margin = Cm(margin)
    section.right_margin = Cm(margin)


contact_info = document.add_paragraph()
contact_info.alignment = WD_ALIGN_PARAGRAPH.DISTRIBUTE

def format_size_and_font(format_obj):
    format_obj.font.size = Pt(12)

def format_alignment(para_obj):
    para_obj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para_obj.left_indent = Inches(0.5)
    return para_obj

#Contact information
address_obj = contact_info.add_run(address + " , ")
phone_obj = contact_info.add_run(phone + " , ")
mail_obj = contact_info.add_run(email)
website_obj = contact_info.add_run("\n" + "Website:" + website)
github_obj = contact_info.add_run("\n" + "Github:" + github)

document.add_heading("_" * 50, 6)

# datetime
today = datetime.datetime.today()
date_obj =  document.add_paragraph().add_run(today.strftime('%d, %b %Y'))


re_obj = document.add_paragraph()
re_obj.add_run("Re: " + position)

dear_obj = document.add_paragraph()
dear_obj.add_run("Dear Hiring Manager,")


#Open paragraph
openpara_obj = document.add_paragraph()
openpara_obj.add_run(open_para.format(position, company_name))
openpara_obj = format_alignment(openpara_obj)


#First coop
first_coop_obj = document.add_paragraph()
first_coop_obj.add_run(first_coop)
first_coop_obj = format_alignment(first_coop_obj)

#Second coop
second_coop_obj = document.add_paragraph()
second_coop_obj.add_run(second_coop)
second_coop_obj = format_alignment(second_coop_obj)


#Activies
activities_obj = document.add_paragraph()
activities_obj = format_alignment(activities_obj)

for activity in activities:
    activities_obj.add_run(activity + ". ")

#Close
close_obj = document.add_paragraph()
close_obj.add_run(closing.format(company_name, phone, email, website ))

# document.add_page_break()

document.save('news.docx')