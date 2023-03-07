set PATH=C:\Projects\git\qtest-via-api-compiled;%PATH%

rm -rf logs/*
rm -rf TestCaseWise/*
rm -rf mapping/*
python -W ignore run.py
pause