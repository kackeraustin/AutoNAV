using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;

namespace AutoNAVSearchSets
{
    public partial class MainWindow : Window
    {
        private Document _doc;

        public MainWindow()
        {
            InitializeComponent();
            _doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
        }

        private void BtnCreateDisciplineSets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCreateDisciplineSets.IsEnabled = false;
                UpdateStatus("Creating discipline search sets...", "#3498DB");

                CreateDisciplineSearchSets();

                UpdateStatus("Discipline search sets created successfully!", "#27AE60");
                MessageBox.Show(
                    "Discipline search sets have been created successfully.\n\nCheck the 'Sets' panel under '1. DISCIPLINES' folder.",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error creating discipline search sets", "#E74C3C");
                MessageBox.Show(
                    "Error creating discipline search sets:\n\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                btnCreateDisciplineSets.IsEnabled = true;
            }
        }

        private void BtnCreateElementSets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCreateElementSets.IsEnabled = false;
                UpdateStatus("Creating element category search sets...", "#3498DB");

                string selectedParameter = GetSelectedParameter();
                CreateElementCategorySearchSets(selectedParameter);

                UpdateStatus("Element category search sets created successfully!", "#27AE60");
                MessageBox.Show(
                    "Element category search sets have been created successfully.\n\nCheck the 'Sets' panel under '2. CLASH SETS' folder.",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error creating element search sets", "#E74C3C");
                MessageBox.Show(
                    "Error creating element search sets:\n\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                btnCreateElementSets.IsEnabled = true;
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void UpdateStatus(string message, string colorHex)
        {
            txtStatus.Text = message;
            txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorHex));
        }

        private string GetSelectedParameter()
        {
            if (rbCategory.IsChecked == true) return "Category";
            if (rbSystemName.IsChecked == true) return "SystemName";
            if (rbSystemClassification.IsChecked == true) return "SystemClassification";
            if (rbWorkset.IsChecked == true) return "Workset";
            if (rbFamilyType.IsChecked == true) return "FamilyType";
            return "Category";
        }

        private void CreateDisciplineSearchSets()
        {
            string[] disciplinePatterns = new string[]
            {
                "_ARCH_", "_STRC_", "_MEP_", "_MECH_", "_ELEC_", "_PLUM_",
                "_HVAC_", "_FIRE_", "_CIVIL_", "_SITE_", "_LAND_"
            };

            Dictionary<string, bool> disciplinesFound = new Dictionary<string, bool>();

            foreach (ModelItem item in _doc.Models.RootItems)
            {
                string itemName = item.DisplayName;

                foreach (string pattern in disciplinePatterns)
                {
                    if (itemName.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string disciplineName = pattern.Trim('_');
                        if (!disciplinesFound.ContainsKey(disciplineName))
                        {
                            disciplinesFound[disciplineName] = true;
                        }
                        break;
                    }
                }
            }

            if (disciplinesFound.Count == 0)
            {
                throw new Exception("No discipline patterns found in model files.\n\nPlease ensure your files contain discipline identifiers like _ARCH_, _STRC_, _MEP_, etc.");
            }

            // Create "1. DISCIPLINES" folder
            FolderItem disciplinesFolder = new FolderItem();
            disciplinesFolder.DisplayName = "1. DISCIPLINES";

            foreach (string discipline in disciplinesFound.Keys)
            {
                // Create search with wildcard pattern
                Search search = new Search();
                search.Selection.SelectAll();

                string wildcardPattern = "*_" + discipline + "_*";
                SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Item", "Name")
                    .DisplayStringWildcard(wildcardPattern);

                search.SearchConditions.Add(condition);
                search.Locations = SearchLocations.DescendantsAndSelf;

                // Create selection set
                SelectionSet selSet = new SelectionSet(search);
                selSet.DisplayName = discipline;

                disciplinesFolder.Children.Add(selSet);
            }

            // Add folder to document
            _doc.SelectionSets.AddCopy(disciplinesFolder);
        }

        private void CreateElementCategorySearchSets(string parameterType)
        {
            // Find the "1. DISCIPLINES" folder
            FolderItem disciplinesFolder = null;
            foreach (SavedItem item in _doc.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup && item.DisplayName == "1. DISCIPLINES")
                {
                    disciplinesFolder = (FolderItem)item;
                    break;
                }
            }

            if (disciplinesFolder == null)
            {
                throw new Exception("Please create discipline search sets first (Function 1) before creating element category search sets.");
            }

            // Get all discipline search sets
            Dictionary<string, SelectionSet> disciplineSets = new Dictionary<string, SelectionSet>();
            foreach (SavedItem discItem in disciplinesFolder.Children)
            {
                if (!discItem.IsGroup)
                {
                    SelectionSet set = (SelectionSet)discItem;
                    disciplineSets[set.DisplayName] = set;
                }
            }

            if (disciplineSets.Count == 0)
            {
                throw new Exception("No discipline search sets found. Please run Function 1 first.");
            }

            // Create "2. CLASH SETS" folder
            FolderItem clashSetsFolder = new FolderItem();
            clashSetsFolder.DisplayName = "2. CLASH SETS";

            int totalCreated = 0;

            // For each discipline
            foreach (var discipline in disciplineSets)
            {
                string disciplineName = discipline.Key;
                SelectionSet disciplineSet = discipline.Value;

                // Create folder for this discipline
                FolderItem disciplineFolder = new FolderItem();
                disciplineFolder.DisplayName = disciplineName;

                // Discover unique values for this property within this discipline
                var propertyValues = DiscoverPropertyValues(disciplineSet, parameterType);

                if (propertyValues.Count == 0)
                {
                    continue;
                }

                // Create search sets for each property value
                foreach (string value in propertyValues)
                {
                    try
                    {
                        Search search = new Search();

                        // Set the search to use the discipline search set as the source
                        ModelItemCollection baseItems = disciplineSet.Search.FindAll(_doc, false);
                        search.Selection.CopyFrom(baseItems);

                        // Add the property condition
                        SearchCondition condition = CreatePropertyCondition(parameterType, value);
                        if (condition != null)
                        {
                            search.SearchConditions.Add(condition);
                            search.Locations = SearchLocations.DescendantsAndSelf;

                            SelectionSet propSet = new SelectionSet(search);
                            propSet.DisplayName = value;

                            disciplineFolder.Children.Add(propSet);
                            totalCreated++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error creating {disciplineName}\\{value}: {ex.Message}");
                    }
                }

                // Add discipline folder to clash sets folder
                if (disciplineFolder.Children.Count > 0)
                {
                    clashSetsFolder.Children.Add(disciplineFolder);
                }
            }

            // Add clash sets folder to document
            if (clashSetsFolder.Children.Count > 0)
            {
                _doc.SelectionSets.AddCopy(clashSetsFolder);
            }

            if (totalCreated == 0)
            {
                throw new Exception("No element categories found in the discipline search sets.");
            }
        }

        private List<string> DiscoverPropertyValues(SelectionSet disciplineSet, string propertyType)
        {
            var values = new HashSet<string>();

            // Get all items in the discipline set - this gets the root file items
            ModelItemCollection rootItems = disciplineSet.Search.FindAll(_doc, false);

            if (rootItems.Count == 0)
            {
                return new List<string>();
            }

            // Map property type to category and property names
            string categoryName, propertyName;
            GetPropertyMapping(propertyType, out categoryName, out propertyName);

            // Start from the entire model hierarchy to find all descendants
            // This ensures we capture categories nested at any depth
            foreach (ModelItem rootItem in rootItems)
            {
                // Recursively scan all descendants starting from root
                ScanItemForPropertyValues(rootItem, categoryName, propertyName, values);
            }

            return values.OrderBy(v => v).ToList();
        }

        private void ScanItemForPropertyValues(ModelItem item, string categoryName, string propertyName, HashSet<string> values)
        {
            // Check current item for the property
            try
            {
                foreach (PropertyCategory cat in item.PropertyCategories)
                {
                    if (cat.DisplayName == categoryName || cat.Name == categoryName)
                    {
                        foreach (DataProperty prop in cat.Properties)
                        {
                            if (prop.DisplayName == propertyName || prop.Name == propertyName)
                            {
                                string val = prop.Value.ToDisplayString();
                                if (!string.IsNullOrWhiteSpace(val))
                                {
                                    values.Add(val);
                                }
                                break;
                            }
                        }
                        break;
                    }
                }
            }
            catch { }

            // Recursively scan all children
            if (item.Children != null)
            {
                foreach (ModelItem child in item.Children)
                {
                    ScanItemForPropertyValues(child, categoryName, propertyName, values);
                }
            }
        }

        private void GetPropertyMapping(string propertyType, out string categoryName, out string propertyName)
        {
            switch (propertyType)
            {
                case "Category":
                    categoryName = "Element";
                    propertyName = "Category";
                    break;
                case "SystemName":
                    categoryName = "Element";
                    propertyName = "System Name";
                    break;
                case "SystemClassification":
                    categoryName = "Element";
                    propertyName = "System Classification";
                    break;
                case "Workset":
                    categoryName = "Element";
                    propertyName = "Workset";
                    break;
                case "FamilyType":
                    categoryName = "Element";
                    propertyName = "Type";
                    break;
                default:
                    categoryName = "Element";
                    propertyName = "Category";
                    break;
            }
        }

        private SearchCondition CreatePropertyCondition(string propertyType, string value)
        {
            string categoryName, propertyName;
            GetPropertyMapping(propertyType, out categoryName, out propertyName);

            SearchCondition condition = SearchCondition.HasPropertyByDisplayName(categoryName, propertyName);
            condition = condition.EqualValue(VariantData.FromDisplayString(value));
            return condition;
        }

        private IEnumerable<ModelItem> GetAllDescendants(ModelItem item)
        {
            yield return item;

            if (item.Children != null)
            {
                foreach (ModelItem child in item.Children)
                {
                    foreach (ModelItem descendant in GetAllDescendants(child))
                    {
                        yield return descendant;
                    }
                }
            }
        }
    }
}
