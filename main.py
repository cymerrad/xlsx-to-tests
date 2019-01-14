#!/usr/bin/env python3
import openpyxl as op
from pathlib import PosixPath
import json

_default_file = "ResolverTests.xlsx"
_template_content1 = '''const data = [
  {test_data}
];

describe("{test_name}", () => {{
  const subject = {{}} as any; // IMPLEMENT ME

  it.each(data)("{each_message_format}", async ({all_arguments}) => {{
    const actualOutput = await subject({test_input});
    expect(actualOutput).toBe({test_output});

    // IMPLEMENT ME
  }});
}});
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


def createTestFileContents(sheet, test_subject, test_message='testing %o'):
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
    }


def writeToTestFile(filepath, content):
    with open(filepath, "w") as fp:
        fp.write(content)


def absoluteFileLocation(base):
    output_dir = PosixPath("dist")  # todo: configurable
    if not output_dir.is_dir():
        output_dir.mkdir()
    return (output_dir / PosixPath(base + ".spec.ts")).absolute()


def main():
    wb = op.load_workbook(_default_file)
    for sheetname in wb.sheetnames:
        content_dict = createTestFileContents(wb[sheetname], sheetname)
        content = _template_content1.format(**content_dict)
        writeToTestFile(absoluteFileLocation(sheetname), content)


if __name__ == "__main__":
    main()
