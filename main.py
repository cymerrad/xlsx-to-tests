#!/usr/bin/env python3
import openpyxl as op
from pathlib import PosixPath
import json

_default_file = "ResolverTests.xlsx"
_template_content1 = '''const data = [
	{test_data}
];

describe("{test_name}", () => {
	const subject = {{}} as any; // IMPLEMENT ME

  it.each(data)({each_message_format}, async ({all_arguments}) => {
    const output = await subject({test_input});
    expect(output).toBe({test_output});

		// IMPLEMENT ME
  });
});
'''

# reserved keywords from sheet
_r_input = "input"
_r_output = "output"
_r_comment = "comment"
_r_implemented = "implemented"
_reserved = [_r_input, _r_input, _r_comment, _r_implemented]

_context_key = "context"


def generateTestData(sheet):
    keys_ord = [str(cell.value).lower()
                for cell in sheet[1]]  # lower-case strings
    context_keys = [x for x in keys_ord]  # hard copy
    for k in _reserved:
        context_keys.remove(k)  # remove all the keys we know _always_ exist

    rows = []
    for row in sheet:
        new_row = {}
        for ind, cell in enumerate(row):
            new_row[keys_ord[ind]] = cell.value

        rows.append(new_row)

    # first object looks like: {"input":"input", ... } so it's trash
    rows = rows[1:]
    # rows should contain now objects of key-value pairs: {"input":"some string", ...}

    test_data = []
    for row in rows:
        new_row = {}
        new_row[_r_input] = row[_r_input]
        new_row[_r_output] = row[_r_output]
        new_unique = {}
        for k in context_keys:
            new_unique[k] = row[k]

        new_row[_context_key] = json.dumps(new_unique)

        test_data.append(row)

    # test_data should contain objects of three key-value pairs
    # keys of: "input", "output" and "context"
    data_lines = ["[\"{input}\", \"{context}\", {}}]".format(
        **datum) for datum in test_data]

    return (",\n".join(data_lines), keys_ord)


def createTestFile(filepath, sheet, test_subject, test_message='testing %o'):
    data = generateTestData(sheet)
    test_data = data[0]
    all_args = ",".join(data[1])
    input_args = ",".join(all_args[:-1])
    output_arg = str(all_args[-1])

    with open(filepath, "w") as fp:
        fp.write(_template_content1.format(**{
            "test_data": test_data,
            "test_name": test_subject,
            "each_message_format": test_message,
            "all_arguments": all_args,
            "test_input": input_args,
            "test_output": output_arg,
        }))


def absoluteFileLocation(base):
    return (PosixPath("dist") / PosixPath(base + ".spec.ts")).absolute()


def main():
    wb = op.load_workbook(_default_file)
    for sheetname in wb.sheetnames:
        try:
            createTestFile(absoluteFileLocation(
                sheetname), wb[sheetname], sheetname)
        except Exception as e:
            print("Error occured", e)


if __name__ == "__main__":
    main()
