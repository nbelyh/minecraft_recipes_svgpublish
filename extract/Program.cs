using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace extract
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    // using System.Windows.Forms;
    using Visio = Microsoft.Office.Interop.Visio;

    internal class Program
    {
        class Ingridient
        {
            public string Name { get; set; }
            public string Quantity { get; set; }
            public string Type { get; set; }
        }

        class Recipe
        {
            public bool Used { get; set; }
            public Visio.Shape Shape { get; set; }
            public List<Ingridient> Ingridients { get; set; }
            public string ImageUrl { get; set; }
        }

        /// <summary>
        /// A simple command
        /// </summary>
        /// 
        static void Generate(Visio.Page page)
        {
            var inventory = new Dictionary<string, Recipe>();

            ReadRecipes(inventory, @"C:\Projects\minecraft_recipes_svgpublish\data\Recipes.csv", "craft");
            ReadRecipes(inventory, @"C:\Projects\minecraft_recipes_svgpublish\data\Furnace.csv", "Furnace");

            var layerCraft = page.Layers.Add("craft");
            var layerFurnace = page.Layers.Add("furnace");

            foreach (var kvp in inventory)
            {
                foreach (var ingridient in kvp.Value.Ingridients)
                {
                    if (inventory.TryGetValue(ingridient.Name, out var found))
                    {
                        found.Used = true;
                    }
                }
            }

            var x = 0.0;
            foreach (var kvp in inventory)
            {
                if (kvp.Value.Used || kvp.Value.Ingridients.Count > 0)
                {
                    var shape = page.DrawRectangle(x, 0, x + 4, 3);
                    shape.Text = kvp.Key;

                    var name = kvp.Key;
                    AddProperty(shape, "Item", name);

                    var ingridients = string.Join(", ", kvp.Value.Ingridients.Select(i => $"{i.Name} ({i.Quantity})").ToList());
                    AddProperty(shape, "Ingridients", ingridients);

                    var recipeImage = kvp.Value.ImageUrl;
                    AddProperty(shape, "RecipeImage", recipeImage);

                    kvp.Value.Shape = shape;
                    x += 5;
                }
            }

            foreach (var kvp in inventory)
            {
                foreach (var ingridient in kvp.Value.Ingridients)
                {
                    if (inventory.TryGetValue(ingridient.Name, out var found))
                    {
                        found.Shape.AutoConnect(kvp.Value.Shape, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        var connector = page.Shapes[page.Shapes.Count];
                        AddProperty(connector, "Type", ingridient.Type);
                        AddProperty(connector, "Quantity", ingridient.Quantity);
                        page.Layers[ingridient.Type].Add(connector, 0);
                    }
                }
            }
        }

        private static void ReadRecipes(Dictionary<string, Recipe> inventory, string fileName, string type)
        {
            var lines = File.ReadAllLines(fileName);
            foreach (var line in lines)
            {
                var items = line.Split(';');

                var name = items[0];
                var ingridientQuantities = items[1]?.Split(',')?.ToList();
                var ingridientNames = items[2]?.Split(',')?.ToList();
                var imageUrl = items[3];

                var si = new Recipe
                {
                    Ingridients = new List<Ingridient>(),
                    ImageUrl = imageUrl,
                    Used = false
                };

                for (var i = 0; i < (ingridientNames?.Count ?? 0); ++i)
                {
                    var ingridientName = ingridientNames[i];
                    var ingridientQuantity = ingridientQuantities[i];

                    si.Ingridients.Add(new Ingridient
                    {
                        Name = ingridientName,
                        Quantity = ingridientQuantity,
                        Type = type
                    });

                    if (!inventory.ContainsKey(ingridientName))
                    {
                        inventory.Add(ingridientName, new Recipe
                        {
                            Ingridients = new List<Ingridient>(),
                            Used = false
                        });
                    }
                }

                inventory[name] = si;
            }
        }

        private static void AddProperty(Visio.Shape shape, string property, string name)
        {
            if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] == 0)
                shape.AddSection((short)Visio.VisSectionIndices.visSectionProp);

            var count = shape.Section[(short)Visio.VisSectionIndices.visSectionProp].Count;
            shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionProp, property, 0);
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, count, (short)Visio.VisCellIndices.visCustPropsValue].FormulaU = $"\"{name}\"";
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, count, (short)Visio.VisCellIndices.visCustPropsLabel].FormulaU = $"\"{property}\"";
        }

        static void Main(string[] args)
        {
            var app = new Visio.Application();
            var doc = app.Documents.Add("");
            var page = doc.Pages.Cast<Visio.Page>().First();

            Generate(page);

            doc.SaveAs(@"C:\Projects\minecraft_recipes_svgpublish\data\minecraft_recipes_svgpublish2.vsdx");
            app.Quit();
        }

    }
}
