#!/usr/bin/env python3
import openpyxl as op
from pathlib import PosixPath
import json
import argparse

_default_file = "ResolverTests.xlsx"
_template_content_file = '''const data = [
  {test_data}
];

describe("{test_name}", () => {{
  const subject = {{}} as any; // IMPLEMENT ME

  it.each(data)("{each_message_format}", {test_async}({all_arguments}) => {{
    const actualOutput = await subject({test_input});
    expect(actualOutput).toBe({test_output});

    // IMPLEMENT ME
  }});
}});
'''

_template_content_stdout = '''\n###### {test_name} ######
  {test_data}
'''

# reserved keywords from sheet
_r_input = "input"
_r_output = "output"
_r_comment = "comment"
_r_implemented = "implemented"
_reserved = [_r_input, _r_output, _r_comment, _r_implemented]

_context_key = "context"


def generateTestData(sheet):
    keys_ord = [str(cell.value).lower()
                for cell in sheet[1]]  # lower-case strings
    max_row = len(keys_ord) - 1
    context_keys = [x for x in keys_ord]  # hard copy
    for k in _reserved:
        try:
            # remove all the keys we know _always_ exist
            context_keys.remove(k)
        except ValueError:
            pass

    no_context = False
    if len(context_keys) == 0:
        no_context = True

    rows = []
    for row in sheet:
        new_row = {}
        for ind, cell in enumerate(row):
            if cell.value is None and ind != max_row:  # empty value and non-comment
                break
            new_row[keys_ord[ind]] = cell.value
        else:
            rows.append(new_row)

    # first object looks like: {"input":"input", ... } so it's trash
    rows = rows[1:]
    # rows should contain now objects of key-value pairs: {"input":"some string", ...}

    test_data = []
    for row in rows:
        new_row = {}
        new_row[_r_input] = row[_r_input]
        new_row[_r_output] = row[_r_output]

        new_context = {}
        for k in context_keys:
            # clear the value from single and double quotations (json.dumps adds its own)
            new_context[k] = str(row[k]).lstrip("\"'").rstrip("\"'")
        new_row[_context_key] = json.dumps(new_context)

        test_data.append(new_row)

    # test_data should contain objects of three or two key-value pairs
    # keys of: "input", "output" and "context"; "context" is optional
    datum_string = "[{input}, {output}]" if no_context else "[{input}, {context}, {output}]"
    datum_names = [_r_input, _r_output] if no_context else [
        _r_input, _context_key, _r_output]
    data_lines = [datum_string.format(
        **datum) for datum in test_data]

    return (",\n  ".join(data_lines), datum_names)


def createTestFileContents(sheet, test_subject, test_message='testing %o', test_async=True, **kwargs):
    data = generateTestData(sheet)
    test_data = data[0]
    all_args = ", ".join(data[1])
    input_args = ", ".join(data[1][:-1])
    output_arg = str(data[1][-1])

    return {
        "test_data": test_data,
        "test_name": test_subject,
        "each_message_format": test_message,
        "all_arguments": all_args,
        "test_input": input_args,
        "test_output": output_arg,
        "test_async": "async " if test_async else "",
    }


def writeToTestFile(filepath, content):
    with open(filepath, "w") as fp:
        fp.write(content)


def absoluteFileLocation(output_dir, base):
    out_dir = PosixPath(output_dir)  # todo: configurable
    if not out_dir.is_dir():
        out_dir.mkdir()
    return (out_dir / PosixPath(base + ".spec.ts")).absolute()


def main(input_file, output_dir, only_data=False, **kwargs):
    wb = op.load_workbook(input_file)

    for sheetname in wb.sheetnames:
        content_dict = createTestFileContents(
            wb[sheetname], sheetname, **kwargs)
        if not only_data:
            content = _template_content_file.format(**content_dict)
            writeToTestFile(absoluteFileLocation(
                output_dir, sheetname), content)
        else:
            content = _template_content_stdout.format(**content_dict)
            print(content)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="input M$ Excel file; default is ResolverTests.xlsx",
                        type=str)
    parser.add_argument("--message",
                        help="message template for the test case", type=str)
    parser.add_argument(
        "--output", help="directory in which to dump the tests; default is \"dist\"", type=str, default="dist")
    parser.add_argument(
        "--not_async", help="test case will not be an asynchronous lambda", action="store_true")
    parser.add_argument(
        "--only_data", help="print to stdout only data", action="store_true")

    args = parser.parse_args()
    func_args = {}
    func_args["input_file"] = getattr(args, "file")
    func_args["output_dir"] = getattr(args, "output")
    if getattr(args, "message"):
        func_args["test_message"] = getattr(args, "message")
    if getattr(args, "not_async"):
        func_args["test_async"] = False  # tricky
    if getattr(args, "only_data"):
        func_args["only_data"] = True

    main(**func_args)
