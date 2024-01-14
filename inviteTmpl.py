#link : https://youtu.be/ZAVHbDB5yBQ

from docxtpl import DocxTemplate


personName = ['Sam', 'Yuha', 'Namduni']

for x, p in enumerate(personName):

    inviteDoc = DocxTemplate('Template/inviteTmpl.docx')
    content = {
        'todayStr' : '23-12-22',
        'recipientName' : p,
        'evntDtStr' : '24-10-22',
        'venueStr' : 'Mutwal',
        'senderName' : 'Godfrri'
    }

    inviteDoc.render(content)
    inviteDoc.save('Output/inviteTmpl_Output_{0}.docx'.format(x))