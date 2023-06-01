using YamlDotNet.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

var yamlConfigPath = "./test.yaml";
var pptxOutputPath = "./output/todo-makethisavariable.pptx";


var yamlConfig = await File.ReadAllTextAsync(yamlConfigPath);
var config = new Deserializer().Deserialize<PowerPointConfig>(yamlConfig);

PowerPointGenerator.Generate(config, pptxOutputPath);

public static class PowerPointGenerator
{
    public static void Generate(PowerPointConfig config, string outputPath)
    {
        using var presentation = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);
        var presentationPart = presentation.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        var slideIdList = new SlideIdList();

        foreach (var slideConfig in config.Slides)
        {
            var slidePart = AddSlide(presentationPart, slideConfig);
            var slideId = new SlideId { Id = (UInt32Value) (101U + slideIdList.ChildElements.Count), RelationshipId = presentationPart.GetIdOfPart(slidePart) };
            slideIdList.Append(slideId);
        }

        presentationPart.Presentation.Append(slideIdList);
        presentationPart.Presentation.Save();
    }

    private static SlidePart AddSlide(PresentationPart presentationPart, PowerPointConfig.SlideConfig slideConfig)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        var slide = new Slide(new CommonSlideData(new ShapeTree()));
        var titleShape = CreateTitleShape(slideConfig.Title);
        var contentShape = CreateContentShape(slideConfig.Content);
        
        if (!string.IsNullOrEmpty(slideConfig.ImagePath))
        {
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using var stream = new FileStream(slideConfig.ImagePath, FileMode.Open);
            imagePart.FeedData(stream);
            
            var imageShape = CreateImageShape(slidePart, imagePart);
            slide.CommonSlideData.ShapeTree.AppendChild(imageShape);
        }

        slide.CommonSlideData.ShapeTree.AppendChild(titleShape);
        slide.CommonSlideData.ShapeTree.AppendChild(contentShape);
        slidePart.Slide = slide;
        return slidePart;
    }


    private static Shape CreateTitleShape(string title)
    {
        var shape = new Shape();
        var textBody = new TextBody();
        var paragraph = new DocumentFormat.OpenXml.Drawing.Paragraph();
        var run = new DocumentFormat.OpenXml.Drawing.Run();
        var text = new Text(title);
        
        run.Append(text);
        paragraph.Append(run);
        textBody.Append(paragraph);
        shape.Append(textBody);
        
        return shape;
    }

    private static Shape CreateContentShape(string content)
    {
        var shape = new Shape();
        var textBody = new TextBody();
        var paragraph = new DocumentFormat.OpenXml.Drawing.Paragraph();
        var run = new DocumentFormat.OpenXml.Drawing.Run();
        var text = new Text(content);
        
        run.Append(text);
        paragraph.Append(run);
        textBody.Append(paragraph);
        shape.Append(textBody);
        
        return shape;
    }

    private static Picture CreateImageShape(SlidePart slidePart, ImagePart imagePart)
    {
        var picture = new Picture();
        var nonVisualPictureProperties = new NonVisualPictureProperties();
        var blipFill = new BlipFill();
        var blip = new DocumentFormat.OpenXml.Drawing.Blip {Embed = slidePart.GetIdOfPart(imagePart)};
        var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
        var fillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

        stretch.Append(fillRectangle);
        blipFill.Append(blip);
        blipFill.Append(stretch);
        picture.Append(nonVisualPictureProperties);
        picture.Append(blipFill);

        return picture;
    }
}


public class PowerPointConfig
{
    public List<SlideConfig> Slides { get; set; }

    public class SlideConfig
    {
        public string Title { get; set; }
        public string Content { get; set; }
        public string ImagePath { get; set; }
    }
}
