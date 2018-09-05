import os
import argparse
import datetime
import errno
import subprocess
import shutil
import json


def main(job_title, company, username=None, password=None, url=None, lan = None):

    synk = False
    if any(v is not None for v in [username, password, url]):
        if any(v is None for v in [username, password, url]):
            print(f'\nOops something went wrong.. You did not input username, password and url.. ALL are required for synking..\n'
                  f'username  -> {True if username else False} \n'
                  f'password  -> {True if password else False}\n'
                  f'url       -> {True if url else False}\n')
            exit(1)
        synk = True

    while lan is not 'DA' and lan is not 'EN':
        lan = decide_lan()

    source = os.path.join('src', 'Danish' if lan is 'DA' else 'English')
    destination = os.path.join('ansoegnings_superfolder', company)

    if synk:
        latex_latest = get_latest(source)
    else:
        latex_latest = get_local_latest(os.path.join(source, 'latex'))

    application(source, job_title, company, destination)

    copy(latex_latest, os.path.join(destination, 'tmp', 'latex'))
    preamble(os.path.join(destination,'tmp','latex','preamble.tex'), company)



def preamble(source, company):
    with open(source, 'r') as f:
        filedate = f.read().replace('COMPANYPLACEHOLDER', company)

    with open(source, 'w') as f:
        f.write(filedate)


def copy(source, destination):
    try:
        print(source)
        print(destination)
        shutil.copytree(source, destination)
    except OSError as e:
        # If the error was caused because the source wasn't a directory
        if e.errno == errno.ENOTDIR:
            print('hmm')
            shutil.copy(source, destination)
        else:
            print('Directory not copied. Error: %s' % e)



def get_local_latest(source):
    local_latest = '2000-01-01 00:00:00'
    folders = [ item for item in os.listdir(source) if os.path.isdir(os.path.join(source, item)) ]
    for folder in folders:
        if datetime.datetime.fromisoformat(folder) > datetime.datetime.fromisoformat(str(local_latest)):
            local_latest = datetime.datetime.fromisoformat(folder)
    return os.path.join(source, str(local_latest))


def get_latest(path):
    print('NOT SUPPORTED YET!!!!!!!!!!1!')
    exit(0)
    path = os.path.join(path + str(datetime.datetime.today()))
    try:
        os.makedirs(path)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise
    return path


def application(source, jobtitle, company, destination):

    try:
        os.makedirs(os.path.join(destination, 'tmp'))
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise

    source = os.path.join(source, 'Application.rtf')
    new_source = os.path.join(destination, 'tmp', 'application.rtf')

    with open(source, 'r') as input:
        with open(new_source, 'w') as output:
            output.write(input.read().replace('JOBTITLE', jobtitle).replace('COMPANY', company))

    convert_to_pdf(new_source, os.path.join(destination, 'folder'))


def decide_lan():
    lan =  input('What language? [DA] or EN: ')
    if len(lan) == 0:
        lan = 'DA'
    elif lan is not 'DA' and lan is not 'EN':
        print('Plz '+ lan +' is not supported')
    return lan


def convert_to_pdf(source, destination):
    try:
        os.makedirs(destination)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise
    bash = f'libreoffice --headless --invisible --norestore --convert-to pdf --outdir {os.path.abspath(destination)} {os.path.abspath(source)}'

    subprocess.check_output(bash.split())

ap = argparse.ArgumentParser(description="Application Generator", epilog="This requires that you have some pretty presice instructions. See sourcecode..")

ap.add_argument('-j', '--jobtitle', help='The jobtitle', type=str, required=True)
ap.add_argument('-c', '--company', help='The jobtitle', type=str, required=True)
ap.add_argument('-l', '--language', help='What language do you want the output in?', type=str)
ap.add_argument('-n', '--username', help='A username and password is required for synking \w sharelatex', default=None)
ap.add_argument('-p', '--password', help='A username and password is required for synking \w sharelatex', default=None)
ap.add_argument('-u', '--url', help='Rightclick on download Source, copy url, paste here', default=None)

if __name__ == '__main__':
    args = ap.parse_args()
    main(args.jobtitle, args.company, args.username, args.password, args.url, args.language)