using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.ServiceModel;
using VisioApi = Microsoft.Office.Interop.Visio;

namespace Microsoft.Crm.Sdk.Samples
{
	/// <summary>
	/// Create a Visio diagram detailing relationships between Microsoft CRM entities.
	///
	/// First, this sample reads in all the entity names. It then creates a visio object for
	/// the entity and all of the entities related to the entity, and links them together.
	/// Finally,it saves the file to disk.
	/// </summary>
	public class DiagramBuilder
	{
        #region Class Level Members

        // Specify which language code to use in the sample. If you are using a language
        // other than US English, you will need to modify this value accordingly.
        // See http://msdn.microsoft.com/en-us/library/0h88fahh.aspx
        //public const int _languageCode = 1033; // English
        public const int _languageCode = 2052; // Chinese

        private VisioApi.Application _application;
        private VisioApi.Document _document;
        private RetrieveMetadataChangesResponse _metadataResponse;
		private ArrayList _processedRelationships;

		private const double X_POS1 = 0;
		private const double Y_POS1 = 0;
		private const double X_POS2 = 1.75;
		private const double Y_POS2 = 0.6;

		const double SHDW_PATTERN = 0;
		const double BEGIN_ARROW_MANY = 29;
		const double BEGIN_ARROW = 0;
		const double END_ARROW = 29;
		const double LINE_COLOR_MANY = 10;
		const double LINE_COLOR = 8;
		const double LINE_PATTERN_MANY = 2;
		const double LINE_PATTERN = 1;
		const string LINE_WEIGHT = "2pt";
		const double ROUNDING = 0.0625;
		const double HEIGHT = 0.25;
		const short NAME_CHARACTER_SIZE = 12;
		const short VISIO_SECTION_OJBECT_INDEX = 1;

        #endregion Class Level Members

        public DiagramBuilder()
		{
			_processedRelationships = new ArrayList(128);
		}
        				
		/// <summary>
		/// Main entry point for the application.
		/// </summary>
		/// <param name="CmdArgs">Entities to place on the diagram</param>
		public static int Main(string[] args)
		{
			String filename = String.Empty;
			VisioApi.Application application;
			VisioApi.Document document;
			DiagramBuilder builder = new DiagramBuilder();

            try
            {
                // Load Visio and create a new document.
                application = new VisioApi.Application();
                application.Visible = false; // Not showing the UI increases rendering speed
                document = application.Documents.Add(String.Empty);

                builder._application = application;
                builder._document = document;

                // Load the metadata.
                Console.WriteLine("Loading Metadata {0} ...", DateTime.Now.ToLongTimeString());
                RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest()
                {
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships,
                    RetrieveAsIfPublished = true
                };

                var response = builder.RetrieveMetadata();
                Console.WriteLine("Metadata Loaded {0} ...", DateTime.Now.ToLongTimeString());
                builder._metadataResponse = response;

                // Diagram all entities if given no command-line parameters, otherwise diagram
                // those entered as command-line parameters.
                if (args.Length < 1)
                {
                    ArrayList entities = new ArrayList();

                    foreach (EntityMetadata entity in response.EntityMetadata)
                    {
                        entities.Add(entity.LogicalName);
                    }

                    builder.BuildDiagram((string[])entities.ToArray(typeof(string)), "All Entities");
                    filename = "AllEntities.vsd";
                }
                else
                {
                    builder.BuildDiagram(args, String.Join(", ", args));
                    filename = String.Concat(args[0], ".vsd");
                }

                // Save the diagram in the current directory using the name of the first
                // entity argument or "AllEntities" if none were given. Close the Visio application. 
                document.SaveAs(Directory.GetCurrentDirectory() + "\\" + filename);
                application.Quit();
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine("Timestamp: {0}", ex.Detail.Timestamp);
                Console.WriteLine("Code: {0}", ex.Detail.ErrorCode);
                Console.WriteLine("Message: {0}", ex.Detail.Message);
                Console.WriteLine("Plugin Trace: {0}", ex.Detail.TraceText);
                Console.WriteLine("Inner Fault: {0}",
                    null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
            }
            catch (System.TimeoutException ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine("Message: {0}", ex.Message);
                Console.WriteLine("Stack Trace: {0}", ex.StackTrace);
                Console.WriteLine("Inner Fault: {0}",
                    null == ex.InnerException.Message ? "No Inner Fault" : ex.InnerException.Message);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine(ex.Message);

                // Display the details of the inner exception.
                if (ex.InnerException != null)
                {
                    Console.WriteLine(ex.InnerException.Message);

                    FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe
                        = ex.InnerException
                        as FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>;
                    if (fe != null)
                    {
                        Console.WriteLine("Timestamp: {0}", fe.Detail.Timestamp);
                        Console.WriteLine("Code: {0}", fe.Detail.ErrorCode);
                        Console.WriteLine("Message: {0}", fe.Detail.Message);
                        Console.WriteLine("Plugin Trace: {0}", fe.Detail.TraceText);
                        Console.WriteLine("Inner Fault: {0}",
                            null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
                    }
                }
            }
            // Additional exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
            // SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.

            finally
            {
                //Console.WriteLine("Rendering complete.");
                Console.WriteLine("Rendering complete.  Press any key to continue.");
                Console.ReadLine();
            }
					
			return 0;
		}

		/// <summary>
		/// Create a new page in a Visio file showing all the direct entity relationships participated in
		/// by the passed-in array of entities.
		/// </summary>
		/// <param name="entities">Core entities for the diagram</param>
		/// <param name="pageTitle">Page title</param>
		private void BuildDiagram(string[] entities, string pageTitle)
		{
			// Get the default page of our new document
			VisioApi.Page page = _document.Pages[1];
			page.Name = pageTitle;
            int i = 1;

			// Get the metadata for each passed-in entity, draw it, and draw its relationships.
			foreach (string entityName in entities)
			{
                Console.Write("Processing entity {1}/{2} {3}: {0}", entityName, i, entities.Length, DateTime.Now.ToLongTimeString());

				EntityMetadata entity = GetEntityMetadata(entityName);

				// Create a Visio rectangle shape.
				VisioApi.Shape rect;
				
				try
				{
					// There is no "Get Try", so we have to rely on an exception to tell us it does not exists
					// We have to skip some entities because they may have already been added by relationships of another entity
					rect = page.Shapes.get_ItemU(entity.LogicalName);
				}
				catch(System.Runtime.InteropServices.COMException)
				{
                    rect = DrawEntityRectangle(page, entity.LogicalName, null == entity.DisplayName.UserLocalizedLabel ? "" : entity.DisplayName.UserLocalizedLabel.Label, entity.OwnershipType.Value);
					Console.Write('.'); // Show progress
				}
				
				// Draw all relationships TO this entity.
				DrawRelationships(entity, rect, entity.ManyToManyRelationships, false);
				Console.Write('.'); // Show progress
				DrawRelationships(entity, rect, entity.ManyToOneRelationships, false);
				
				// Draw all relationshipos FROM this entity
				DrawRelationships(entity, rect, entity.OneToManyRelationships, true);
				Console.WriteLine('.'); // Show progress

                i++;
			}

			// Arrange the shapes to fit the page.
			page.Layout();
			page.ResizeToFitContents();
		}

        private RetrieveMetadataChangesResponse RetrieveMetadata()
        {
            // Connect to the Organization service. 
            // The using statement assures that the service proxy will be properly disposed.
            using (var service = new OrganizationService("Crm"))
            {
                //A filter expression to limit entities returned to non-intersect, user-owned entities not found in the list of excluded entities.
                MetadataFilterExpression EntityFilter = new MetadataFilterExpression(LogicalOperator.And);
                var entities = ConfigurationManager.AppSettings["Entities"];
                if (!string.IsNullOrWhiteSpace(entities))
                {
                    string[] includedEntities = entities.ToLowerInvariant().Replace(" ", string.Empty).Split(new string[]{ "," }, StringSplitOptions.RemoveEmptyEntries);
                    EntityFilter.Conditions.Add(new MetadataConditionExpression("LogicalName", MetadataConditionOperator.In, includedEntities));
                }                
                EntityFilter.Conditions.Add(new MetadataConditionExpression("IsValidForAdvancedFind", MetadataConditionOperator.Equals, true));

                //A properties expression to limit the properties to be included with entities
                MetadataPropertiesExpression EntityProperties = new MetadataPropertiesExpression(
                    "OwnershipType",
                    "DisplayName",
                    "Attributes",
                    "ManyToManyRelationships",
                    "ManyToOneRelationships",
                    "OneToManyRelationships",
                    "PrimaryIdAttribute")
                {
                    AllProperties = false
                };

                //A label query expression to limit the labels returned to only those for the user's preferred language
                LabelQueryExpression labelQuery = new LabelQueryExpression();
                labelQuery.FilterLanguages.Add(_languageCode);

                //A condition expresson to return optionset attributes
                MetadataConditionExpression[] optionsetAttributeTypes = new MetadataConditionExpression[] { 
                    new MetadataConditionExpression("AttributeType", MetadataConditionOperator.Equals, AttributeTypeCode.Uniqueidentifier),
                    new MetadataConditionExpression("AttributeType", MetadataConditionOperator.Equals, AttributeTypeCode.Customer),
                    new MetadataConditionExpression("AttributeType", MetadataConditionOperator.Equals, AttributeTypeCode.Lookup),
                    new MetadataConditionExpression("AttributeType", MetadataConditionOperator.Equals, AttributeTypeCode.Owner)
                };

                //A filter expression to apply the optionsetAttributeTypes condition expression
                MetadataFilterExpression AttributeFilter = new MetadataFilterExpression(LogicalOperator.Or);
                AttributeFilter.Conditions.AddRange(optionsetAttributeTypes);

                //A Properties expression to limit the properties to be included with attributes
                MetadataPropertiesExpression AttributeProperties = new MetadataPropertiesExpression() { AllProperties = false };
                AttributeProperties.PropertyNames.Add("DisplayName");

                MetadataPropertiesExpression RelationshipProperties = new MetadataPropertiesExpression(
                    "Entity1LogicalName",
                    "Entity2LogicalName",
                    "RelationshipType",
                    "MetadataId",
                    "ReferencingEntity",
                    "ReferencedEntity",
                    "ReferencingAttribute",
                    "ReferencedAttribute")
                {
                    AllProperties = false
                };

                //An entity query expression to combine the filter expressions and property expressions for the query.
                EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
                {
                    Criteria = EntityFilter,
                    Properties = EntityProperties,
                    AttributeQuery = new AttributeQueryExpression()
                    {
                        Criteria = AttributeFilter,
                        Properties = AttributeProperties
                    },
                    RelationshipQuery = new RelationshipQueryExpression()
                    {
                        Properties = RelationshipProperties
                    },
                    LabelQuery = labelQuery
                };

                RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
                {
                    Query = entityQueryExpression
                };

                return (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);
            }
        }

		/// <summary>
		/// Draw an "Entity" Rectangle
		/// </summary>
		/// <param name="page">The Page on which to draw</param>
		/// <param name="entityName">The name of the entity</param>
		/// <param name="ownership">The ownership type of the entity</param>
		/// <returns>The newly drawn rectangle</returns>
		private VisioApi.Shape DrawEntityRectangle(VisioApi.Page page, string entityName, string displayName, OwnershipTypes ownership)
		{
			VisioApi.Shape rect = page.DrawRectangle(X_POS1, Y_POS1, X_POS2, Y_POS2);
			rect.Name = entityName;
            rect.Text = displayName + " ";

			// Determine the shape fill color based on entity ownership.
			string fillColor;

			switch (ownership)
			{
				case OwnershipTypes.BusinessOwned:
					fillColor = "RGB(255,202,176)"; // Light orange
					break;
				case OwnershipTypes.OrganizationOwned:
					fillColor = "RGB(255,255,176)"; // Light yellow
					break;
				case OwnershipTypes.UserOwned:
					fillColor = "RGB(204,255,204)"; // Light green
					break;
				default:
					fillColor = "RGB(255,255,255)"; // White
					break;
			}

			// Set the fill color, placement properties, and line weight of the shape.
			rect.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowMisc, (short)VisioApi.VisCellIndices.visLOFlags).FormulaU = ((int)VisioApi.VisCellVals.visLOFlagsPlacable).ToString();
			rect.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowFill, (short)VisioApi.VisCellIndices.visFillForegnd).FormulaU = fillColor;

			// Update the style of the entity name
			VisioApi.Characters characters = rect.Characters;
			characters.Begin = 0;
            characters.End = displayName.Length;
			characters.set_CharProps((short)VisioApi.VisCellIndices.visCharacterStyle, (short)VisioApi.VisCellVals.visBold);
			characters.set_CharProps((short)VisioApi.VisCellIndices.visCharacterColor, (short)VisioApi.VisDefaultColors.visDarkBlue);
			characters.set_CharProps((short)VisioApi.VisCellIndices.visCharacterSize, NAME_CHARACTER_SIZE);

			return rect;
		}
		
		/// <summary>
		/// Draw a directional, dynamic connector between two entities, representing an entity relationship.
		/// </summary>
		/// <param name="shapeFrom">Shape initiating the relationship</param>
		/// <param name="shapeTo">Shape referenced by the relationship</param>
		/// <param name="isManyToMany">Whether or not it is a many-to-many entity relationship</param>
		private void DrawDirectionalDynamicConnector(VisioApi.Shape shapeFrom, VisioApi.Shape shapeTo, bool isManyToMany)
		{
			// Add a dynamic connector to the page.
			VisioApi.Shape connectorShape = shapeFrom.ContainingPage.Drop(_application.ConnectorToolDataObject, 0.0, 0.0);

			// Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
			connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowFill, (short)VisioApi.VisCellIndices.visFillShdwPattern).ResultIU = SHDW_PATTERN;
			connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowLine, (short)VisioApi.VisCellIndices.visLineBeginArrow).ResultIU = isManyToMany ? BEGIN_ARROW_MANY : BEGIN_ARROW;
			connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowLine, (short)VisioApi.VisCellIndices.visLineEndArrow).ResultIU = END_ARROW;
			connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowLine, (short)VisioApi.VisCellIndices.visLineColor).ResultIU = isManyToMany ? LINE_COLOR_MANY : LINE_COLOR;
			connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowLine, (short)VisioApi.VisCellIndices.visLinePattern).ResultIU = isManyToMany ? LINE_PATTERN : LINE_PATTERN;
			connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowFill, (short)VisioApi.VisCellIndices.visLineRounding).ResultIU = ROUNDING;

			// Connect the starting point.
			VisioApi.Cell cellBeginX = connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowXForm1D, (short)VisioApi.VisCellIndices.vis1DBeginX);
			cellBeginX.GlueTo(shapeFrom.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowXFormOut, (short)VisioApi.VisCellIndices.visXFormPinX));

			// Connect the ending point.
			VisioApi.Cell cellEndX = connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowXForm1D, (short)VisioApi.VisCellIndices.vis1DEndX);
			cellEndX.GlueTo(shapeTo.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowXFormOut, (short)VisioApi.VisCellIndices.visXFormPinX));
		}

		/// <summary>
		/// Retrieves an entity from the local copy of CRM Metadata
		/// </summary>
		/// <param name="entityName">The name of the entity to find</param>
		/// <returns>NULL if the entity was not found, otherwise the entity's metadata</returns>
		private EntityMetadata GetEntityMetadata(string entityName)
		{
			foreach (EntityMetadata md in _metadataResponse.EntityMetadata)
			{
				if (md.LogicalName == entityName)
				{
					return md;
				}
			}

			return null;
		}

		/// <summary>
		/// Retrieves an attribute from an EntityMetadata object
		/// </summary>
		/// <param name="entity">The entity metadata that contains the attribute</param>
		/// <param name="attributeName">The name of the attribute to find</param>
		/// <returns>NULL if the attribute was not found, otherwise the attribute's metadata</returns>
		private AttributeMetadata GetAttributeMetadata(EntityMetadata entity, string attributeName)
		{
			foreach (AttributeMetadata attrib in entity.Attributes)
			{
				if (attrib.LogicalName == attributeName)
				{
					return attrib;
				}
			}

			return null;
		}

        /// <summary>
        /// Draw on a Visio page the entity relationships defined in the passed-in relationship collection.
        /// </summary>
        /// <param name="entity">Core entity</param>
        /// <param name="rect">Shape representing the core entity</param>
        /// <param name="relationshipCollection">Collection of entity relationships to draw</param>
        /// <param name="areReferencingRelationships">Whether or not the core entity is the referencing entity in the relationship</param>
        private void DrawRelationships(EntityMetadata entity, VisioApi.Shape rect, RelationshipMetadataBase[] relationshipCollection, bool areReferencingRelationships)
        {
            ManyToManyRelationshipMetadata currentManyToManyRelationship = null;
            OneToManyRelationshipMetadata currentOneToManyRelationship = null;
            EntityMetadata entity2 = null;
            AttributeMetadata attribute2 = null;
            AttributeMetadata attribute = null;
            Guid metadataID = Guid.NewGuid();
            bool isManyToMany = false;

            // Draw each relationship in the relationship collection.
            foreach (RelationshipMetadataBase entityRelationship in relationshipCollection)
            {
                entity2 = null;

                if (entityRelationship is ManyToManyRelationshipMetadata)
                {
                    isManyToMany = true;
                    currentManyToManyRelationship = entityRelationship as ManyToManyRelationshipMetadata;
                    // The entity passed in is not necessarily the originator of this relationship.
                    if (String.Compare(entity.LogicalName, currentManyToManyRelationship.Entity1LogicalName, true) != 0)
                    {
                        entity2 = GetEntityMetadata(currentManyToManyRelationship.Entity1LogicalName);
                    }
                    else
                    {
                        entity2 = GetEntityMetadata(currentManyToManyRelationship.Entity2LogicalName);
                    }
                    if (entity2 == null)
                    {
                        continue;
                    }
                    attribute2 = GetAttributeMetadata(entity2, entity2.PrimaryIdAttribute);
                    attribute = GetAttributeMetadata(entity, entity.PrimaryIdAttribute);
                    metadataID = currentManyToManyRelationship.MetadataId.Value;
                }
                else if (entityRelationship is OneToManyRelationshipMetadata)
                {
                    isManyToMany = false;
                    currentOneToManyRelationship = entityRelationship as OneToManyRelationshipMetadata;
                    entity2 = GetEntityMetadata(areReferencingRelationships ? currentOneToManyRelationship.ReferencingEntity : currentOneToManyRelationship.ReferencedEntity);
                    if (entity2 == null)
                    {
                        continue;
                    }
                    attribute2 = GetAttributeMetadata(entity2, areReferencingRelationships ? currentOneToManyRelationship.ReferencingAttribute : currentOneToManyRelationship.ReferencedAttribute);
                    attribute = GetAttributeMetadata(entity, areReferencingRelationships ? currentOneToManyRelationship.ReferencedAttribute : currentOneToManyRelationship.ReferencingAttribute);
                    metadataID = currentOneToManyRelationship.MetadataId.Value;
                }
                // Verify relationship is either ManyToManyMetadata or OneToManyMetadata
                if (entity2 != null)
                {
                    if (_processedRelationships.Contains(metadataID))
                    {
                        // Skip relationships we have already drawn
                        continue;
                    }
                    else
                    {
                        // Record we are drawing this relationship
                        _processedRelationships.Add(metadataID);

                        // Define convenience variables based upon the direction of referencing with respect to the core entity.
                        VisioApi.Shape rect2;


                        // Do not draw relationships involving the entity itself, SystemUser, BusinessUnit,
                        // or those that are intentionally excluded.
                        if (String.Compare(entity2.LogicalName, "systemuser", true) != 0 &&
                            String.Compare(entity2.LogicalName, "businessunit", true) != 0 &&
                            String.Compare(entity2.LogicalName, rect.Name, true) != 0 &&
                            String.Compare(entity.LogicalName, "systemuser", true) != 0 &&
                            String.Compare(entity.LogicalName, "businessunit", true) != 0)
                        {
                            // Either find or create a shape that represents this secondary entity, and add the name of
                            // the involved attribute to the shape's text.
                            try
                            {
                                rect2 = rect.ContainingPage.Shapes.get_ItemU(entity2.LogicalName);

                                if (null != attribute2.DisplayName.UserLocalizedLabel && rect2.Text.IndexOf(attribute2.DisplayName.UserLocalizedLabel.Label) == -1)
                                {
                                    rect2.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowXFormOut, (short)VisioApi.VisCellIndices.visXFormHeight).ResultIU += 0.25;
                                    rect2.Text += "\n" + attribute2.DisplayName.UserLocalizedLabel.Label;

                                    // If the attribute is a primary key for the entity, append a [PK] label to the attribute name to indicate this.
                                    if (String.Compare(entity2.PrimaryIdAttribute, attribute2.LogicalName) == 0)
                                    {
                                        rect2.Text += "  [PK]";
                                    }
                                }
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                rect2 = DrawEntityRectangle(rect.ContainingPage, entity2.LogicalName, null == entity2.DisplayName.UserLocalizedLabel ? "" : entity2.DisplayName.UserLocalizedLabel.Label, entity2.OwnershipType.Value);
                                if (attribute2.DisplayName.UserLocalizedLabel != null)
                                {
                                    rect2.Text += "\n" + attribute2.DisplayName.UserLocalizedLabel.Label;
                                }

                                // If the attribute is a primary key for the entity, append a [PK] label to the attribute name to indicate so.
                                if (String.Compare(entity2.PrimaryIdAttribute, attribute2.LogicalName) == 0)
                                {
                                    rect2.Text += "  [PK]";
                                }
                            }

                            // Add the name of the involved attribute to the core entity's text, if not already present.
                            if (null != attribute.DisplayName.UserLocalizedLabel && rect.Text.IndexOf(attribute.DisplayName.UserLocalizedLabel.Label) == -1)
                            {
                                rect.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)VisioApi.VisRowIndices.visRowXFormOut, (short)VisioApi.VisCellIndices.visXFormHeight).ResultIU += HEIGHT;
                                rect.Text += "\n" + attribute.DisplayName.UserLocalizedLabel.Label;

                                // If the attribute is a primary key for the entity, append a [PK] label to the attribute name to indicate so.
                                if (String.Compare(entity.PrimaryIdAttribute, attribute.LogicalName) == 0)
                                {
                                    rect.Text += "  [PK]";
                                }
                            }

                            // Draw the directional, dynamic connector between the two entity shapes.
                            if (areReferencingRelationships)
                            {
                                DrawDirectionalDynamicConnector(rect, rect2, isManyToMany);
                            }
                            else
                            {
                                DrawDirectionalDynamicConnector(rect2, rect, isManyToMany);
                            }
                        }
                        else
                        {
                            Debug.WriteLine(String.Format("<{0} - {1}> not drawn.", rect.Name, entity2.LogicalName), "Relationship");
                        }
                    }
                }
            }
        }
    }
}
