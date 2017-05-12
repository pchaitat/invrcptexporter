# invrcptexporter
a very simple tool to export invoice and receipt prepared with libreoffice calc into pdf file

This program exports invoice and receipt in PDF format.

virtualenv has to be created with option `--system-site-packages` as it
needs to use system site-packages which has uno module installed by
`$ sudo apt-get install python3-uno`, on Ubuntu.

To open LibreOffice with socket open, use the command line below:

```
libreoffice \
  --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
```

INSTALL

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

    TO_BE_CONTINUED

  - configure `~/.invrcptexporterrc`:

      ```
      $ cd ~
      $ cp /your/project/path/source/invrcptexporterrc .invrcptexporterrc
      ```

    configure settings as you need inside this rc file, the most
    important one is SCRIPT_PATH which is the absolute path to 'source'
    directory
