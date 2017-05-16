# invrcptexporter

a tool to export invoice and receipt from libreoffice calc into pdf file

virtualenv has to be created with option `--system-site-packages` as it
needs to use system site-packages which has uno module installed by
`$ sudo apt-get install python3-uno`, on Ubuntu.  (Unfortunately, uno
module is not available from pip.)

The code is based on Christopher5106's tutorial (many thanks to him):

  ```
  ('http://christopher5106.github.io/office/2015/12/06/openoffice-',
   'libreoffice-automate-your-office-tasks-with-python-macros.html')
  ```

## How it works

- we open LibreOffice with socket open
- we use Python to send commands to LibreOffice through the socket (you
  should use iptables to block incoming connection on the specified TCP
  port for security reason)

To open LibreOffice with socket open, use the command line below:

```
libreoffice \
  --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
```

## INSTALL

  - On Ubuntu 14.04

  - install python3-uno as a system-wide package as this package is not
    available via pip:

      ```
      $ sudo apt-get install python3-uno
      ```

  - create virtualenv (probably, with virtualenvwrapper) with
    `--system-site-pacakges`:

      ```
      $ mkvirtualenv --system-site-packages invrcptexporter
      ```

  - install required packages from pip:

      ```
      $ pip install -r requirement.txt
      ```

  - install poppler-utils from apt as we have to use pdftotext in our
    tests:

      ```
      $ sudo apt-get install poppler-utils
      ```

    I have tried a few other alternatives such as extractText() from
    PyPDF2, it doesn't work for many cases

  - configure `~/.invrcptexporterrc`:

      ```
      $ cd ~
      $ cp /your/project/path/source/invrcptexporterrc .invrcptexporterrc
      ```

    adjust settings in this rc file, the most important one is
    SCRIPT_PATH which is the absolute path to 'source' directory

  - run the test:

      ```
      $ /your/project/path/source/test.sh
      ```

    notes: test.sh is a hack (you can see details from comments inside
    test.sh ), each time you run it, you have to manually close (exit or
    ctrl+q) the opened libreoffice ui so that the test will continue to
    run
