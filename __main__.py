"""Push/pull OOo macro code between files and a live OOo instance."""
from contextlib import contextmanager
from operator import methodcaller

def parse_arg(args):
    if len(args) == 4:
        return args[1:]

add_document_arg = methodcaller('add_argument', 'document',
    help="The name of a document which is open in the running OO.o app.  "
            "E.G. 'Untitled 1'")
add_macro_arg = methodcaller('add_argument', 'macro',
    help="The name of the macro library or routine.  "
            "E.G. 'Standard.Module1.main'")
def add_source_file_arg(parser, mode='read'):
    from sys import stdin, stdout
    from argparse import FileType
    filemode, default, default_name = {
        'read': ('r', stdin, 'standard in'),
        'write': ('w', stdout, 'standard out')}[mode]
    parser.add_argument('source_file',
        type=FileType(filemode), nargs='?', default=default,
        help="The name of the file which contains the source code. "
             "Defaults to {0}.".format(default_name))

def parse_args():
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command')

    push_command = commands.add_parser(
        'push', help="Push code into the given module.")
    add_document_arg(push_command)
    add_macro_arg(push_command)
    add_source_file_arg(push_command, 'read')

    pull_command = commands.add_parser(
        'pull', help="Pull code from the given module.")
    add_document_arg(pull_command)
    add_macro_arg(pull_command)
    add_source_file_arg(pull_command, 'write')

    invoke_command = commands.add_parser(
        'invoke', help="Invoke the given macro.")
    add_document_arg(invoke_command)
    add_macro_arg(invoke_command)

    return parser.parse_args()

def take_action(options):
    """Takes the action prescribed by `options.command`.

    See `parse_args` for option details.
    """
    from exchange import Exchange

    action = getattr(Exchange(), options.command)

    arg_attrs = ['document', 'macro']

    if options.command == 'push':
        arg_attrs.append('source_file')

    result = action(*(getattr(options, attr) for attr in arg_attrs))

    if options.command == 'pull':
        options.source_file.writelines(line + "\n" for line in result)

@contextmanager
def close_source(options):
    """Ensure that the file gets properly closed on program termination."""
    yield options
    if hasattr(options, 'source_file'):
        options.source_file.close()

def main():
    with close_source(parse_args()) as options:
        take_action(options)

# Clear out module-level imports
del contextmanager, methodcaller

if __name__ == '__main__':
    exit(main())
