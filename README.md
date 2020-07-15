# docx-playground

At the moment this provides functionality to convert from `.docx` into both `.html` and `.json` files. It's not perfect. Garbage in, expect garbage out.

Add or remove files from the `files/` directory for them to be converted. It will ignore hidden files and non-`.docx` files.

## getting started
1. run `nvm use` (get nvm if you don't have it)
2. run `npm install` (get npm if you don't have it)
3. run `npm run start`
4. results are outputted to `output/*.html` and `output/*.json`