
## Currently only capable of generating my RTU MIREA frontend reports

### Requirements

`pip install imgkit`<br>
`pip install python-docx`

### Usage
`python gen.py <path_to_directory>`

I was quite lazy to make the tool usable so to generate a report you should have:
+ `template.docx` document with title and some styles in script local directory
+ all your html task sorted in directories with names like `task<n>`, each directory has `index.html`
+ `task.html` file in target directory, containing html with task name in `<h3></h3>` and task text as text

example structure:

```
project
|-- task.html
|-- task1
    |-- index.html
|-- task2
...
```

see https://github.com/Frovu/mirea-frontend for full target example
