using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using log4net;

namespace MigrationApiDemo
{
    public class ManifestPackage
    {
        private readonly SharePointMigrationTarget _target;
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public ManifestPackage(SharePointMigrationTarget sharePointMigrationTarget)
        {
            _target = sharePointMigrationTarget;
        }

        public IEnumerable<MigrationPackageFile> GetManifestPackageFiles(IEnumerable<SourceFile> sourceFiles)
        {
            Log.Debug("Generating manifest package");
            var result = new[]
            {
                GetExportSettingsXml(),
                GetLookupListMapXml(),
                GetManifestXml(sourceFiles),
                GetRequirementsXml(),
                GetRootObjectMapXml(),
                GetSystemDataXml(),
                GetUserGroupXml(),
                GetViewFormsListXml()
            };

            Log.Debug($"Generated manifest package containing {result.Length} files, total size: {result.Select(x => x.Contents.Length).Sum() / 1024.0 / 1024.0:0.00}mb");

            return result;
        }

        private MigrationPackageFile GetExportSettingsXml()
        {
            var exportSettingsDefaultXml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<ExportSettings SiteUrl=\"http://fileshare/sites/user\" FileLocation=\"C:\\Temp\\0 FilesToUpload\" IncludeSecurity=\"None\" xmlns=\"urn:deployment-exportsettings-schema\" />");
            return new MigrationPackageFile { Filename = "ExportSettings.xml", Contents = exportSettingsDefaultXml };
        }

        private MigrationPackageFile GetLookupListMapXml()
        {
            var lookupListMapDefaultXml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<LookupLists xmlns=\"urn:deployment-lookuplistmap-schema\" />");
            return new MigrationPackageFile { Filename = "LookupListMap.xml", Contents = lookupListMapDefaultXml };
        }

        private MigrationPackageFile GetRequirementsXml()
        {
            var requirementsDefaultXml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<Requirements xmlns=\"urn:deployment-requirements-schema\" />");
            return new MigrationPackageFile { Filename = "Requirements.xml", Contents = requirementsDefaultXml };
        }

        private MigrationPackageFile GetRootObjectMapXml()
        {
            var objectRootMapDefaultXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n";
            objectRootMapDefaultXml += "<RootObjects xmlns=\"urn:deployment-rootobjectmap-schema\">";
            objectRootMapDefaultXml +=
                $"<RootObject Id=\"{_target.DocumentLibraryId}\" Type=\"List\" ParentId=\"{_target.WebId}\" WebUrl=\"{_target.SiteName}\" Url=\"{string.Format($"{_target.SiteName}/{_target.DocumentLibraryName}", _target.SiteName, _target.DocumentLibraryName)}\" IsDependency=\"false\" />";

            return new MigrationPackageFile { Filename = "RootObjectMap.xml", Contents = Encoding.UTF8.GetBytes(objectRootMapDefaultXml) };
        }

        private MigrationPackageFile GetSystemDataXml()
        {
            var systemDataXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                                "<SystemData xmlns=\"urn:deployment-systemdata-schema\">" +
                                "<SchemaVersion Version=\"15.0.0.0\" Build=\"16.0.3111.1200\" DatabaseVersion=\"11552\" SiteVersion=\"15\" ObjectsProcessed=\"106\" />" +
                                "<ManifestFiles>" +
                                "<ManifestFile Name=\"Manifest.xml\" />" +
                                "</ManifestFiles>" +
                                "<SystemObjects>" +
                                "</SystemObjects>" +
                                "<RootWebOnlyLists />" +
                                "</SystemData>";
            return new MigrationPackageFile { Filename = "SystemData.xml", Contents = Encoding.UTF8.GetBytes(systemDataXml) };
        }

        private MigrationPackageFile GetUserGroupXml()
        {
            var userGroupDefaultXml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<UserGroupMap xmlns=\"urn:deployment-usergroupmap-schema\"><Users /><Groups /></UserGroupMap>");
            return new MigrationPackageFile { Filename = "UserGroup.xml", Contents = userGroupDefaultXml };
        }

        private MigrationPackageFile GetViewFormsListXml()
        {
            var viewFormsListDefaultXml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<ViewFormsList xmlns=\"urn:deployment-viewformlist-schema\" />");
            return new MigrationPackageFile { Filename = "ViewFormsList.xml", Contents = viewFormsListDefaultXml };
        }

        private MigrationPackageFile GetManifestXml(IEnumerable<SourceFile> files)
        {
            var webUrl = $"{_target.SiteName}";
            var documentLibraryLocation = $"{webUrl}/{_target.DocumentLibraryName}";
            var subfolderLocation = $"{documentLibraryLocation}/{_target.Subfolder}";

            var rootNode = new SPGenericObjectCollection1();

            var rootfolder = new SPGenericObject
            {
                Id = _target.RootFolderId.ToString(),
                ObjectType = SPObjectType.SPFolder,
                ParentId = _target.RootFolderParentId.ToString(),
                ParentWebId = _target.WebId.ToString(),
                ParentWebUrl = webUrl,
                Url = documentLibraryLocation,
                Item = new SPFolder
                {
                    Id = _target.RootFolderId.ToString(),
                    Url = _target.DocumentLibraryName,
                    Name = _target.DocumentLibraryName,
                    ParentFolderId = _target.RootFolderParentId.ToString(),
                    ParentWebId = _target.WebId.ToString(),
                    ParentWebUrl = webUrl,
                    ContainingDocumentLibrary = _target.DocumentLibraryId.ToString(),
                    TimeCreated = DateTime.Now,
                    TimeLastModified = DateTime.Now,
                    SortBehavior = "1",
                    Properties = null
                }
            };
            rootNode.SPObject.Add(rootfolder);

            var documentLibrary = new SPGenericObject
            {
                Id = _target.DocumentLibraryId.ToString(),
                ObjectType = SPObjectType.SPDocumentLibrary,
                ParentId = _target.WebId.ToString(),
                ParentWebId = _target.WebId.ToString(),
                ParentWebUrl = webUrl,
                Url = documentLibraryLocation,
                Item = new SPDocumentLibrary
                {
                    Id = _target.DocumentLibraryId.ToString(),
                    BaseTemplate = "DocumentLibrary",
                    ParentWebId = _target.WebId.ToString(),
                    ParentWebUrl = webUrl,
                    RootFolderId = _target.RootFolderId.ToString(),
                    RootFolderUrl = documentLibraryLocation,
                    Title = _target.DocumentLibraryName
                }
            };
            rootNode.SPObject.Add(documentLibrary);

            var counter = 0;
            foreach (var file in files)
            {
                counter++;
                var fileId = Guid.NewGuid();

                var spFile = new SPGenericObject
                {
                    Id = fileId.ToString(),
                    ObjectType = SPObjectType.SPFile,
                    ParentId = _target.RootFolderId.ToString(),
                    ParentWebId = _target.WebId.ToString(),
                    ParentWebUrl = webUrl,
                    Url = $"{subfolderLocation}/{file.Filename}",
                    Item = new SPFile
                    {
                        Id = fileId.ToString(),
                        Url = $"{_target.DocumentLibraryName}/{_target.Subfolder}/{file.Filename}",
                        Name = $"{_target.Subfolder}/{file.Filename}",
                        ListItemIntId = counter,
                        ListId = _target.DocumentLibraryId.ToString(),
                        ParentId = _target.RootFolderId.ToString(),
                        ParentWebId = _target.WebId.ToString(),
                        TimeCreated = file.LastModified,
                        TimeLastModified = file.LastModified,
                        Version = "1.0",
                        FileValue = file.Filename,
                        Versions = null,
                        Properties = null,
                        WebParts = null,
                        Personalizations = null,
                        Links = null,
                        EventReceivers = null
                    }
                };
                rootNode.SPObject.Add(spFile);

                var spListItemContainerId = Guid.NewGuid();
                var spListItemContainer = new SPGenericObject
                {
                    Id = spListItemContainerId.ToString(),
                    ObjectType = SPObjectType.SPListItem,
                    ParentId = _target.DocumentLibraryId.ToString(),
                    ParentWebId = _target.WebId.ToString(),
                    ParentWebUrl = webUrl,
                    Url = $"{subfolderLocation}/{file.Filename}",
                    Item = new SPListItem
                    {
                        FileUrl = $"{_target.DocumentLibraryName}/{_target.Subfolder}/{file.Filename}",
                        DocType = ListItemDocType.File,
                        ParentFolderId = _target.RootFolderId.ToString(),
                        Order = counter * 100,
                        Id = spListItemContainerId.ToString(),
                        ParentWebId = _target.WebId.ToString(),
                        ParentListId = _target.DocumentLibraryId.ToString(),
                        Name = $"{_target.Subfolder}/{file.Filename}",
                        DirName = "/sites/user/Documents", //todo Migration: are we allways storing in documents directory?
                        IntId = counter,
                        DocId = fileId.ToString(),
                        Version = "1.0",
                        TimeLastModified = file.LastModified,
                        TimeCreated = file.LastModified,
                        ModerationStatus = SPModerationStatusType.Approved
                    }
                };

                var spfields = new SPFieldCollection();
                foreach (var fileProp in file.Properties)
                {
                    var spfield = new SPField();

                    var isMultiValueTaxField = false; //todo
                    var isTaxonomyField = false; //todo

                    if (isMultiValueTaxField)
                    {
                        //todo
                        //spfield.Name = [TaxHiddenFieldName];
                        //spfield.Value = "[guid-of-hidden-field]|[text-value];[guid-of-hidden-field]|[text-value2];";
                        //spfield.Type = "Note"; 
                    }
                    else if (isTaxonomyField)
                    {
                        //todo
                        //spfield.Name = [TaxHiddenFieldName];
                        //spfield.Value = [Value] + "|" + [TaxHiddenFieldValue];
                        //spfield.Type = "Note"; 
                    }
                    else
                    {
                        spfield.Name = fileProp.Key;
                        spfield.Value = fileProp.Value;
                        spfield.Type = "Text";
                    }
                    spfields.Field.Add(spfield);
                }

                var titleSpField = new SPField();
                titleSpField.Name = "Title";
                titleSpField.Value = file.Title;
                titleSpField.Type = "Text";
                spfields.Field.Add(titleSpField);

                ((SPListItem)spListItemContainer.Item).Items.Add(spfields);
                rootNode.SPObject.Add(spListItemContainer);
            }
            var serializer = new XmlSerializer(typeof(SPGenericObjectCollection1));

            var settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            //settings.OmitXmlDeclaration = false;

            using (var memoryStream = new MemoryStream())
            using (var xmlWriter = XmlWriter.Create(memoryStream, settings))
            {
                serializer.Serialize(xmlWriter, rootNode);
                return new MigrationPackageFile
                {
                    Contents = memoryStream.ToArray(),
                    Filename = "Manifest.xml"
                };
            }
        }
    }
}