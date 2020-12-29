import argparse

parser = argparse.ArgumentParser(description='Process schedule and synchronise with Google Calendar.')
parser.add_argument('--dry-run', dest='dry', action='store_const',
                    const=True, default=False,
                    help='dry-run, do not upload results.')
