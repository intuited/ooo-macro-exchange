"""Push/pull OOo macro code between files and a live OOo instance."""
from contextlib import contextmanager
from operator import methodcaller
from functools import partial

from library import IllegalMacroNameError


# Parse

def parse_arg(args):
    if len(args) == 4:
        return args[1:]


add_document_arg = methodcaller('add_argument', 'document',
    help="The name of a document which is open in the running OO.o app.  "
         "E.G. 'Untitled 1'")

def add_macro_arg(parser,
    help="The macro library and module, e.g. 'Standard.Module1'", **kwargs):
    return parser.add_argument('macro', help=help, **kwargs)

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


def ArgumentParser():
    """Build a parser for the command line.

    Returns both the parser and the dictionary of subparsers.

    There is one subparser for each command.

    The main reason for returning the commands
    is so that if further validation on them fails,
    the error method of the appropriate parser can be called.
    """
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--host', default='localhost')
    parser.add_argument('-p', '--port', default='2002')

    commands = parser.add_subparsers(dest='command')

    push_command = commands.add_parser(
        'push', help="Push code into the given module.")
    push_command.add_argument('-s', '--save', action='store_true',
        help="Save the document after updating the macro module.")
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
    add_macro_arg(invoke_command,
        help="The fully-qualified macro name, e.g. 'Standard.Module1.main'")

    return parser, commands.choices


# Action

def split_macro_name(command, macro_name):
    """Split `macro_name` into arguments to the `command`.

    Drops all but the first two components for push and pull commands.

    For other commands, returns a single element
    containing the original string.
    """
    if command in ('push', 'pull'):
        args = macro_name.split('.')
        if len(args) < 2:
            raise IllegalMacroNameError(
                "Macro name must contain at least Library.Module parts.")
        return args[:2]
    else:
        # Other commands (just `invoke` for now)
        # take the unparsed macro name.
        return [macro_name]

def take_action(options):
    """Takes the action prescribed by `options.command`.

    The command name is mapped directly to a method of `exchange.Exchange`.

    See `parse_args` for option details.
    """
    from exchange import Exchange

    action = getattr(Exchange(host=options.host, port=options.port),
                     options.command)

    args = [options.document] + split_macro_name(options.command,
                                                 options.macro)
    kwargs = {}

    if options.command == 'push':
        args.append(options.source_file)
        kwargs.update(save=options.save)

    result = action(*args, **kwargs)

    if options.command == 'pull':
        options.source_file.writelines(line + "\n" for line in result)


# Main

@contextmanager
def close_source(options):
    """Ensure that the file gets properly closed on program termination."""
    yield options
    if hasattr(options, 'source_file'):
        options.source_file.close()

def main():
    parser, commands = ArgumentParser()
    with close_source(parser.parse_args()) as options:
        try:
            take_action(options)
        except IllegalMacroNameError as e:
            commands[options.command].error(str(e))


# Clear out module-level imports
del contextmanager, methodcaller

if __name__ == '__main__':
    exit(main())
