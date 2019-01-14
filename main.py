#!/usr/bin/env python3
import openpyxl as op
from pathlib import PosixPath

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


def generateTestData(sheet):
    return ("\n".join([]), [])


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
