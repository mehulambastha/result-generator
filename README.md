# Result Generator

---

### Small node.js program which converts students' result sent by CBSE in .txt format and writes the data into a spreadsheet.

---

> _**NOTE**: You must have node installed in your system to use this program._

#### How to use ?

1. Clone the repository to your system.
2. Place your text file (.txt file) containing the exam results of students in this directory.

   > _The text file must be in the format of the [sample text file in this repository](./SAMPLE_FILE.txt)_
3. Open [result10.js](./resul10.js) file, go to line 63, and change the "SAMPLE_FILE.txt" to the name of your text file.
   > e.g. -<br> change **`fs.createReadStream('SAMPLE_FILE.txt')`** to **`fs.createReadStream('_name of your text file here_')`**
4. Navigate to the this folder in terminal.
5. Run **`node install`** in the terminal.
6. After installation of dependencies is complete, run **`node result10`** in the same directory from terminal.
7. Your spreadsheet will be generated in the same directory.
