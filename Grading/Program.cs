using Aspose.Cells;
using System.Xml.Serialization;

Grading g = new Grading();
g.EstimateAll();
Console.WriteLine("Таблица успешно обновлена");

class Grading
{
    Workbook workbook;
    Worksheet worksheet;
    int numberOfStudents;
    int maxExercises;
    double maxAttendance = 0;
    double avgExercises;
    double avgAttendance;
    int offsetThreshold = 10;
    int pointsForTasks = 60;
    int estimationColumnIndex;
    string filename;
    string outputFilename;
    Style offsetStyle;
    Style failStyle;
    public Grading()
    {
        Init();
        workbook = new Workbook(filename);
        WorksheetCollection collection = workbook.Worksheets;
        worksheet = collection[0];
        offsetStyle = worksheet.Cells[0, 0].GetStyle();
        failStyle = worksheet.Cells[0, 0].GetStyle();
        offsetStyle.SetPatternColor(BackgroundType.Gray50, System.Drawing.Color.FromArgb(255, 0, 255, 0), System.Drawing.Color.FromArgb(255, 0, 255, 0));
        failStyle.SetPatternColor(BackgroundType.Gray50, System.Drawing.Color.FromArgb(255, 255, 0, 0), System.Drawing.Color.FromArgb(255, 255, 0, 0));
        offsetStyle.Font.IsBold = true;
        failStyle.Font.IsBold = true;
    }

    void Init()
    {
        ConfigSettings settings = new ConfigSettings();
        using (StreamReader reader = new StreamReader("config.txt"))
        {
            XmlSerializer serializer = new XmlSerializer(typeof(ConfigSettings));
            settings = (ConfigSettings)serializer.Deserialize(reader);
        }

        filename = settings.filename;
        offsetThreshold = settings.offsetThreshold;
        pointsForTasks = settings.pointsForTasks;
        outputFilename = settings.outputFilename;
    }

    public void EstimateAll()
    {
        AnalyzeData();
        worksheet.Cells[0, worksheet.Cells.MaxDataColumn + 1].PutValue("Итоговая оценка");
        for (int i = 0; i < numberOfStudents; i++)
        {
            EstimateOne(i + 1);
        }
        workbook.Save(outputFilename, SaveFormat.Xlsx);
    }


    void AnalyzeData()
    {
        int rows = worksheet.Cells.MaxDataRow + 1;
        int cols = worksheet.Cells.MaxDataColumn + 1;
        maxExercises = 0;
        numberOfStudents = rows - 1;
        estimationColumnIndex = cols;

        double tmpAttendance = 0;
        double tmpExercises = 0;

        for (int i = 1; i < rows; i++)
        {
            if (double.Parse(worksheet.Cells[i, 2].Value.ToString()) > maxAttendance)
                maxAttendance = double.Parse(worksheet.Cells[i, 2].Value.ToString());
            tmpAttendance += double.Parse(worksheet.Cells[i, 2].Value.ToString());

            int temp = 0;
            for(int j = 3; j < cols; j++)
            {
                if (worksheet.Cells[i, j].Value != null)
                { 
                    tmpExercises++;
                    temp++;
                }
            }
            if(temp > maxExercises)
                maxExercises = temp;
        }
        avgAttendance = tmpAttendance / numberOfStudents;
        avgExercises = tmpExercises / numberOfStudents;
    }
    void EstimateOne(int index)
    {
        int completedExercises = 0;
        for (int i = 3; i < estimationColumnIndex; i++)
        {
            if (worksheet.Cells[index, i].Value != null)
                completedExercises++;
        }
        int grade = CalculateEstimation(completedExercises, double.Parse(worksheet.Cells[index, 2].Value.ToString()), (worksheet.Cells[index, 1].Value != null)? int.Parse(worksheet.Cells[index, 1].Value.ToString()) : 60);
        worksheet.Cells[index, estimationColumnIndex].PutValue(grade);
        if (grade < 60)
            worksheet.Cells[index, estimationColumnIndex].SetStyle(failStyle);
        else
            worksheet.Cells[index, estimationColumnIndex].SetStyle(offsetStyle);
    }

    int CalculateEstimation(int completedExercises, double attendance, int personalEstimation)
    {
        double result;
        result = pointsForTasks * ((double)completedExercises / maxExercises) + (100 - pointsForTasks) * attendance;
        if(result < personalEstimation)
        {
            if (result < 60)
                personalEstimation = 60;
            int delta = personalEstimation - (int)result;
            delta = (int)(delta * (((double)completedExercises / maxExercises) + attendance) / 2);
            result += delta;
        }
        if (result < 60)
            if (60 - result <= offsetThreshold)
                result = 60;
        if (result > 100)
            result = 100;
        return (int)result;
    }

}

[Serializable] public class ConfigSettings
{
    [XmlElement("filename")] public string filename { get; set; }
    [XmlElement("offsetThreshold")] public int offsetThreshold { get; set; }
    [XmlElement("pointsForTasks")] public int pointsForTasks { get; set; }
    [XmlElement("outputFilename")] public string outputFilename { get; set; }
}