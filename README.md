# XLS Redirect Generator

This is a simple node command line tool for generating .htaccess rewrite redirects from any xls/xlsx file. A basic redirect checker is also included and can work off the same file used to generate the redirects.

## Generator Usage

Clone the repository and install the dependencies from your terminal:

```
git clone https://github.com/zefman/xlsredirect.git
cd xlsredirect
npm install
```

Generate the rewrite rules by using the following command, and then answering the questions in your terminal:

```
node index.js
```

The rewrite rules will be generated in a file called output.txt in the same directory. These can then be copied into your .htaccess file.

## Checker Usage

The checker works in an almost identical way to the generator. Start the generator with the following command, and complete the questions.

```
node checker.js
```

Once finished the checker will generated a file called failed.txt in the same directory with details about any redirects that didn't work as expected so you can debug.
