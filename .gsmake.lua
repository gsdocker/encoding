name "github.com/gsdocker/encoding"

plugin "github.com/gsmake/golang"


properties.golang = {
    dependencies = {
        { name = "github.com/gsdocker/gsos"     };
        { name = "github.com/gsdocker/gserrors" };
        { name = "github.com/gsdocker/gsconfig" };
        { name = "github.com/gsdocker/gslogger" };
        { name = "github.com/tealeg/xlsx" };
    };

    tests = {"."}
}
