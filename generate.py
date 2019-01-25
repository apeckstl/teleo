from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.text import PP_ALIGN
import sys

TITLE_LAYOUT = 0
BLANK = 10
FONT = "Calibri"

class LiturgyElem:
    def __init__(self,f_name):
        self.title = ""
        self.lyrics = []
        self.f_name = f_name

    #opens a file and turns it into a LiturgyElem
    def parseLiturgyFile(self,delimiter):
        with open(self.f_name,"r") as file:
            lyrics = file.read()
            lyrics = lyrics.split(delimiter)
            self.title = lyrics[0]
            self.lyrics = lyrics[1:]

#creates a textbox for a song slide, with parameters for placement, size, and boldness
#returns a paragraph!!
def createVerseTextBox(verse_slide,dim,size,bold):
        shapes = verse_slide.shapes
        verse_text = shapes.add_textbox(Inches(dim[0]),Inches(dim[1]),Inches(dim[2]),Inches(dim[3]))
        frame = verse_text.text_frame
        frame.word_wrap = True
        p = frame.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Calibri'
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = RGBColor(255,255,255)
        return p

#method done. Adds a song to the presentation at the end, pulling from a Song object
def addSong(song,prs):
    #add title slide for song
    song_title_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
    title = createVerseTextBox(song_title_slide,(0.375,0.25,9.25,1.25),36,True)
    title.text = song.title
    #put first verse of song on title slide
    p = createVerseTextBox(song_title_slide,(0.5,1.5,9,6),30,False)
    p.text = song.lyrics[0]
    #loop through remaining verses and add slides
    for verse in song.lyrics[1:]:
        verse_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
        p = createVerseTextBox(verse_slide,(0.5,1.5,9,6),30,False)
        p.text = verse

#Done, creates text frame on a slide with given dimensions and adds 2 paragraphs
def createTextBox(slide,dim):
        shapes = slide.shapes
        verse_text = shapes.add_textbox(Inches(dim[0]),Inches(dim[1]),Inches(dim[2]),Inches(dim[3]))
        frame = verse_text.text_frame
        frame.word_wrap = True
        p = frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Calibri'
        p.font.color.rgb = RGBColor(255,255,255)
        p = frame.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Calibri'
        p.font.color.rgb = RGBColor(255,255,255)
        p = frame.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Calibri'
        p.font.color.rgb = RGBColor(255,255,255)
        return frame

def addCallConfSlide(ref,prs,flag):
    #add title slide for service
    first_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
    heading = createTextBox(first_slide,(0.5,0.25,9.5,7.25))
    head = ref.title.split('\n')
    title = heading.paragraphs[0]
    title.font.size = Pt(32)
    title.font.bold = True
    title.text = head[0]
    i = 1
    if flag:
        i = 2
        mid = heading.paragraphs[1]
        mid.font.size = Pt(24)
        mid.text = "Now hear the gracious word of God from:"
    sub = heading.paragraphs[i]
    sub.font.size = Pt(30)
    sub.font.bold = True
    sub.text = head[1]
    #put first verse of song on title slide
    p = createVerseTextBox(first_slide,(0.5,1.5,9,6),28,False)
    p.alignment = PP_ALIGN.LEFT
    p.text = ref.lyrics[0]
    #loop through remaining verses and add slides
    for verse in ref.lyrics[1:]:
        verse_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
        p = createVerseTextBox(verse_slide,(0.5,1,9,6.5),28,False)
        p.alignment = PP_ALIGN.LEFT
        p.text = verse


def addBlank(prs):
    prs.slides.add_slide(prs.slide_layouts[BLANK])

#open template
prs = Presentation("template.pptx")
addBlank(prs)

#welcome slide
welcome_slide = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
title_placeholder = welcome_slide.placeholders[0]
title_placeholder.text = "Welcome to Living Hope Church!"
addBlank(prs)

#first song
song = LiturgyElem(sys.argv[1])
song.parseLiturgyFile("\n\n")
addSong(song,prs)
addBlank(prs)

#call to worship
call = LiturgyElem(sys.argv[2])
call.parseLiturgyFile("$$$\n")
addCallConfSlide(call,prs,False)
addBlank(prs)

#second, third, fourth song
for i in range(3,6):
    print(sys.argv[i])
    song = LiturgyElem(sys.argv[i])
    song.parseLiturgyFile("\n\n")
    addSong(song,prs)
    addBlank(prs)

conf = LiturgyElem(sys.argv[6])
conf.parseLiturgyFile("$$$\n")
addCallConfSlide(conf,prs,False)
addBlank(prs)

#assurance of grace
grace = LiturgyElem(sys.argv[7])
grace.parseLiturgyFile("$$$\n")
addCallConfSlide(grace,prs,True)
addBlank(prs)

#offering
offering_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
heading = createVerseTextBox(offering_slide,(0,0.5,10,1.5),32,True)
heading.text = "Worship through Tithes and Offerings"
p = createVerseTextBox(offering_slide,(0,1.5,10,6),28,False)
with open(sys.argv[8],"r") as file:
    offering = file.read()
p.text = offering
addBlank(prs)

#offering song
song = LiturgyElem(sys.argv[9])
song.parseLiturgyFile("\n\n")
addSong(song,prs)
addBlank(prs)

#Prayers of the people
prayer_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
text = createVerseTextBox(prayer_slide,(0,1.5,10,1.5),30,True)
text.text = "Prayers of the People"
addBlank(prs)

#Announcements
ann_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
text = createVerseTextBox(ann_slide,(0.25,0,10,7.25),28,True)
text.text = "Announcements"
addBlank(prs)

sermon_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
heading = createTextBox(sermon_slide,(0.75,1.5,8.5,4.5))
with open( sys.argv[10],"r") as file:
    name = file.read()
title_parts = name.split("\n")
first_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
title = heading.paragraphs[0]
title.font.size = Pt(30)
title.font.bold = True
title.text = title_parts[0]
mid = heading.paragraphs[1]
mid.font.size = Pt(24)
mid.font.italic = True
mid.font.bold = True
mid.text = title_parts[1]
sub = heading.paragraphs[2]
sub.font.size = Pt(30)
sub.font.bold = True
sub.text = title_parts[2]

#closing song
song = LiturgyElem(sys.argv[11])
song.parseLiturgyFile("\n\n")
addSong(song,prs)
addBlank(prs)

bene_slide = prs.slides.add_slide(prs.slide_layouts[BLANK])
text = createVerseTextBox(bene_slide,(0,0.5,10,1),44,True)
text.text = "Benediction"
rest = createVerseTextBox(bene_slide,(0,1.75,10,3.5),40,False)
rest.text = "A blessing from God’s Word\nas we go out into God’s world."

prs.save("slidesbase.pptx")
