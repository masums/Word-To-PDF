# Word-To-PDF
Word TO PDF Net Core Library 
## How To Use

Create instance of `DocumentService` class, `DocumentService` class is class which responsible for generated pdf document

``` c#
try
{
  DocumentService documentService = new DocumentService("wordtemplate/test.docx");
}
catch (Exception)
{
  throw;
}
```


### Add Data to replace

```c#
  documentService.Data = new DataField()
   {
        Data = new Dictionary<string, string>()
        {
            {  "Image", "Images.PNG" },
            { "Avatar",  "https://pay.google.com/about/static/images/social/og_image.jpg"},
            {  "FirstName", "User 1" }
        },
            Options = new Dictionary<string, Options>()
        {
            { "Avatar", new Options(){  FromUrl = true, PercentageResize = 10 } },
            { "Image", new Options(){  FromUrl = false, Width = 100 } }
        }
    },
```

### Add Object Data to replace

```c#
documentService.DataTable = new List<DataFieldGorup()
{
    {
        new DataFieldGroup()
        {
            Data = new List<User>()
            {
                new User()
                {
                    Nama = "user 1",
                    Email = "user1@gmail.com",
                    Photo = "Image.PNG",
                },
                new User()
                {
                    Nama = "user 2",
                    Email = "user2,cmon",
                    Photo = "test.png",
                }
            },
            Key = "User",
            Options = new Dictionary<string, Options>()
            {
                { "Photo", new Options()
                    {
                        FromUrl = false,
                        Width = 200,
                        Height = 300
                    }
                }
            }
        }
    },
    {
        new DataFieldGroup()
        {
            Data = new List<Pegawai>()
            {
                new Pegawai()
                {
                    Nama = "Pegawai 1",
                    Unit = "Pusintek",
                    Image = "https://pay.google.com/about/static/images/social/og_image.jpg",
                },
                new Pegawai()
                {
                    Nama = "Pegawai 2",
                    Unit = "Pusintek",
                    Image = "https://pbs.twimg.com/profile_images/972154872261853184/RnOg6UyU.jpg",
                }
            },
            Key = "Pegawai",
            Options = new Dictionary<string, Options>()
            {
                { "Image", new Options()
                    {
                        FromUrl = true,
                        Width = 200,
                        Height = 300
                    }
                }
            },
        }
    }
}
```
