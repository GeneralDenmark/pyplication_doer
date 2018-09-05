import os, argparse, datetime, errno, comtypes.client
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
    destination = os.path.join('ansøgnings_superfolder', company)



    if synk:
        get_latest(source)


def preamble(path, jobtitle, company, destination):
    pass


def get_latest(path):
    try:
        os.makedirs(os.path.join(path + str(datetime.datetime.today())))
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise


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
    lan =  input('What language? [DA] or EN:    ')
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

    wd_format_PDF = 17
    out_file = os.path.join(destination, 'ansøgning.pdf')
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(source)
    doc.SaveAs(out_file, FileFormat=wd_format_PDF)
    doc.Close()
    word.Quit()


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