# coding=utf-8
import datetime
import tempfile
from collections import defaultdict
import cherrypy as cherrypy
import os
from cherrypy.lib.static import serve_file
from mailmerge import MailMerge
from wtforms import Form, StringField, BooleanField
import dominate.tags as html
from wtforms.fields.core import SelectField
from wtforms.fields.html5 import DateField
from wtforms.validators import InputRequired

currdir = os.path.dirname(__file__)


cherrypy.config.update({
    'server.socket_host': '0.0.0.0',
    'engine.autoreload.on': False,
    'server.socket_port': 9090 if 'PORT' not in os.environ else int(os.environ['PORT']),
    'tools.sessions.on': True,
    'tools.sessions.timeout': 1000 * 60,
    'tools.check_ssl.on': True,
    'tools.sessions.httponly': True,
    'tools.secureheaders.on': True,
    'tools.encode.on': True,
    'tools.encode.encoding': 'utf-8',
    'tools.gzip.mime_types': ['text/html', 'text/plain', 'text/javascript', 'text/css', 'text/csv', 'application/javascript'],
    'tools.gzip.on': True,
})


class Template(object):

    def __init__(self, filename: str):
        self.template = None
        with open(filename) as f:
            self.template = f.read()

    def render(self, **kwargs) -> str:
        """
        Fills the template with the named arguments, fills in empty string if not provided
        """
        return self.template.format_map(defaultdict(lambda: '', **kwargs))


def secureheaders():
    """
    This enables some header options to avoid spoofing and other attacks
    """
    headers = cherrypy.response.headers
    headers['X-Frame-Options'] = 'DENY'
    headers['X-XSS-Protection'] = '1; mode=block'
    headers['Strict-Transport-Security'] = 'max-age=31536000'  # one year


def check_ssl():
    """
    Switches to ssl if on cloud
    """
    if 'PORT' not in os.environ:  # local
        return
    header = cherrypy.request.headers.get('X-Forwarded-Proto', None)
    # print(cherrypy.request.headers)
    if header is None or header == 'http':
        raise cherrypy.HTTPRedirect(cherrypy.url().replace("http:", "https:"))


# ssl
cherrypy.tools.check_ssl = cherrypy.Tool('before_handler', check_ssl)
cherrypy.tools.secureheaders = cherrypy.Tool('before_finalize', secureheaders, priority=60)


template = os.path.join(currdir, "documents/02 Gründungsurkunde (...) GmbH.docx")


class SimpleForm(Form):

    def html(self, action: str, method: str='GET', enctype: str="multipart/form-data", id_: str='my-form') -> html.html_tag:
        """
        Renders the HTML of the form
        """
        content = html.div(cls='vis-form')

        form_args = dict(action=action, cls='form-horizontal', method=method, enctype=enctype, id=id_)

        form_html = content.add(html.form(**form_args))

        for name, field in self._fields.items():

            row = form_html.add(html.div(cls='form-group row_vertical_offset'))

            label = row.add(html.div(cls='col-sm-4 control-label'))
            label.add_raw_string(str(field.label))

            cell = row.add(html.div(cls='col-sm-8'))
            cell.add_raw_string(field(class_="form-control"))

        # submit button
        button_row = form_html.add(html.div(cls='text-right row_vertical_offset'))
        button_row.add(html.button('Senden', type="submit", id='submit_form_button', cls='btn btn-primary'))

        return content


class OkForm(SimpleForm):
    einverstanden = SelectField(label='Ich bin einverstanden', choices=(('Ja', 'Ja'), ('Nein', 'Nein')), validators=[InputRequired()])


class KantonsForm(SimpleForm):
    kanton = SelectField(label='Kanton', choices=(('Zug', 'Zug'), ('Zürich', 'Zürich')), validators=[InputRequired()])




class NewUserForm(SimpleForm):
    gmbh_name = StringField(label='Name der GmbH')
    sitz = StringField(label='Sitz')
    # notariat = StringField()

    # anrede_bevollmaechtigter = SelectField(choices=(('Herr', 'Herr'), ('Frau', 'Frau')))
    # vorname_bevollmaechtigter = StringField()
    # nachname_bevollmaechtigter = StringField()
    # geburtstag_bevollmaechtigter = DateField(format='%d.%m.%Y')
    # buergerort_bevollmaechtigter = StringField()
    # wohnort_bevollmaechtigter = StringField()

    anrede_gruender = SelectField(label='Anrede', choices=(('Herr', 'Herr'), ('Frau', 'Frau')))
    vorname_gruender = StringField(label='Vorname')
    nachname_gruender = StringField(label='Nachname')
    telefon_gruender = StringField(label='Telefon')
    geburtstag_gruender = DateField(label='Geburtsdatum', format='%d.%m.%Y')
    buergerort_gruender = StringField(label='Bürgerort')
    strasse_gruender = StringField(label='Adresse')
    wohnort_gruender = StringField(label='Postleitzahl und Ort')

    # bank = StringField()

    branche_taetigkeit = StringField(label='Branche/Tätigkeit')
    zweck = StringField(label='Zweck')


class Root:
    def __init__(self):
        self.template = Template(os.path.join(currdir, 'template.html'))

    @cherrypy.expose
    def index(self):
        with html.div() as content:
            html.div('Blabla Irgendwas Disclaimer')
            form = OkForm()
            form.html(action='/step1')
        return self.template.render(content=content)

    @cherrypy.expose
    def step1(self, **kwargs):
        if kwargs.get('einverstanden') != 'Ja':
            return self.template.render(content=html.div('Dann halt nicht'))

        with html.div() as content:
            html.div('Wählen Sie Ihren Kanton')
            form = KantonsForm()
            form.html(action='/step2')
        return self.template.render(content=content)

    @cherrypy.expose
    def step2(self, **kwargs):
        form = NewUserForm()
        return self.template.render(content=form.html(action='/create'))




    @cherrypy.expose
    def create(self, **kwargs):
        kwargs['bank'] = 'ZKB, Abteilung SCBJ3, Postfach, 8010 Zürich'
        kwargs['datum_auftrag'] = datetime.date.today().strftime('%d.%m.%Y')

        document = MailMerge(template)
        print(document.get_merge_fields())
        document.merge(**kwargs)

        # noinspection PyProtectedMember
        default_tmp_dir = tempfile._get_default_tempdir()
        # noinspection PyProtectedMember
        temp_name = next(tempfile._get_candidate_names())
        name = 'generated.docx'
        path = os.path.join(default_tmp_dir, temp_name)
        document.write(path)

        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        return serve_file(os.path.abspath(path), disposition='attachment', content_type=mime_type, name=name)


conf = {
    '/css': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': os.path.join(currdir, 'css'),
        'tools.expires.on': True,
        'tools.expires.secs': 60 * 60 * 24 * 365,
    },
    '/bootstrap': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': os.path.join(currdir, 'bootstrap'),
        'tools.expires.on': True,
        'tools.expires.secs': 60 * 60 * 24 * 365,
    },
    '/images': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': os.path.join(currdir, 'images'),
        'tools.expires.on': True,
        'tools.expires.secs': 60 * 60 * 24 * 365,
    },
    '/js': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': os.path.join(currdir, 'js'),
        'tools.expires.on': True,
        'tools.expires.secs': 60 * 60 * 24 * 365,
    }
}


if __name__ == "__main__":
    cherrypy.quickstart(Root(), '/', config=conf)
