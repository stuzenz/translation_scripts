# Translation Scripts

This github repo is a WIP, that provides machine translation via LLMs

- initially I developed against local models, but have decided to move it towards using APIs
- APIs used are to be used with commercial providers who do not use the data for training models.


## Leadership Board

Currently gemini is leading on the leadership board and is very competitive for pricing, this is the set of models I am focusing on.


## Tooling

- devenv.sh
- git-crypt
- python
- shell scripts

## Standardisation of flags

Standardisation of flags and output is done where possible

```

script input_file -src-lang Japanese --tgt-lang English

```

Additional flags for 

## Scripts

### Visio

```bash
python visio_translator4a.py FUT_Creation\ of\ an\ Organisation.vsdx --target-lang Japanese --batch-size 30
```

### Docx


```bash
python docx_translator4.py --target-lang ja --batch-size 20 --not-to-translate-styles "User Story List,CW PathWay" JL_Q1\ Foundations\ Activity\ Worksheet\ CRM.docx
```