var express = require('express')
var multer = require('multer')
var mongoose = require('mongoose')
var path = require('path')
var bodyParser = require('body-parser')
var empSchema = require('./models/EmpModel')
const app = express()
const reader = require('xlsx')

var storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, './uploads')
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname)
  },
})
var uploads = multer({ storage: storage })

mongoose
  .connect('mongodb://localhost:27017/test', { useNewUrlParser: true })
  .then(() => console.log('Connected'))
  .catch((err) => console.log(err))
app.set('view engine', 'ejs')
app.use(bodyParser.urlencoded({ extended: false }))
app.use(express.static(path.resolve(__dirname, 'public')))
app.get('/', (req, res) => {
  empSchema.find((err, data) => {
    if (err) {
    } else {
      if (data != '') {
        res.render('index', { data: data })
      } else {
        res.render('index', { data: '' })
      }
    }
  })
})

let data = [];

app.post('/add', uploads.single('file-xlsx'), (req, res) => {
    // read xlsx file and upload to mongodb
    const file = reader.readFile(req.file.path)
    let sheet = file.SheetNames;
    for (let i = 0; i < sheet.length; i++) {
        const temp = reader.utils.sheet_to_json(
            file.Sheets[file.SheetNames[i]])
          temp.forEach((res) => {
              data.push(res)
           }
        )
    }
    console.log(data)
    data.forEach((res) => {
        var emp = new empSchema(res)
        emp.save()
    }
    )
    res.redirect('/')

    


})

var port = process.env.PORT || 5555
app.listen(port, () => console.log('App connected on: ' + port))