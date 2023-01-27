from generate_python_docx import generate_docx

with open('static/BRANCH.txt') as f:
    all_branches = f.read()

all_branches = all_branches.splitlines()
for branch in all_branches:
    try:
        generate_docx(branch)
    except Exception as e:
        print(e)


