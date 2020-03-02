using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Foss.FossDoc.ApplicationServer;
using DS;
using Converters = Foss.FossDoc.ApplicationServer.Converters;

namespace document.outgoing.letter
{
	public class BaseDocument
	{
		protected ISession _Session;
		protected OID _ParentFolderID;

		protected TPropertyTag _CreationContainerTag;

		public BaseDocument(ISession session, OID parentFolderID, OID categoryID)
		{
			_Session = session;
			_ParentFolderID = parentFolderID;
			_ID = Converters.OID.NewOID(parentFolderID._base);
			_CategoryID = categoryID;
			_CreationContainerTag = Foss.FossDoc.ExternalModules.BusinessLogic.Schema.Folder.Attributes.Documents.Tag;

			_Props[Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID] = Converters.Properties.ObjectPropertyBuilder.Create(_CategoryID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID);

			if (_GenerateDisplayName)
			{
				//Имя документа сформируем
				string catName = categoryID.ToStringRepresentation();
				var props = _Session.ObjectDataManager.GetProperties(categoryID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName);
				if (props[0].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName))
				{
					catName = props[0].Value.GetstrVal();
				}

				string docName = string.Format("{0} ({1}) {2}", catName, _Session.ObjectDataManager.GetObjectDisplayName(session.AccessControlManager.CurrentUserOID), DateTime.Now);
				_Props[Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName] = Converters.Properties.ObjectPropertyBuilder.Create(docName, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName);
			}
		}

		protected virtual bool _GenerateDisplayName
		{
			get
			{
				return true;
			}
		}

		protected virtual void _BeforeDocumentCreated()
		{
			//тут вы можете добавить свойств в _Props
		}

		public OID Create()
		{
			_BeforeDocumentCreated();

			_Session.ObjectDataManager.CreateObjectWithOID(_ParentFolderID, _CreationContainerTag, _ID, PropertiesForCreation);

			_AfterDocumentCreated();

			return _ID;
		}

		protected Dictionary<TPropertyTag, ObjectProperty> _Props = new Dictionary<TPropertyTag, ObjectProperty>(Converters.PropertyTag.EqualityComparer.Instance);

		public ObjectProperty[] PropertiesForCreation
		{
			get
			{
				return Foss.TemplateLibrary.Sets.GetValueArray<TPropertyTag, ObjectProperty>(_Props);
			}
		}

		protected OID _CategoryID;
		public OID CategoryID
		{
			get
			{
				return _CategoryID;
			}
		}


		protected OID _ID;
		public OID ID
		{
			get
			{
				return _ID;
			}
		}
		
		public Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach[] Attaches
		{
			get;
			set;
		}

		/// <summary>
		/// Если было задано свойство Attaches создает вложения в документе
		/// </summary>
		protected virtual void _CreateAttaches()
		{
			if (Attaches == null || Attaches.Length == 0)
				return;

			foreach (var att in Attaches)
			{
				att.CreateAttach(_Session, ID);
			}
		}

		/// <summary>
		/// Здесь разместите действия, которые нужно выполнить после физического создания документа.
		/// Предполагается что документ уже создан и есть в базе.
		/// </summary>
		protected virtual void _AfterDocumentCreated()
		{
			_CreateAttaches();
		}
		
		public static OID[] Create(ISession session, OID folder, params BaseDocument[] documentsForCreation)
		{
			List<OID> docIDs = new List<OID>();
			List<ObjectProperty[]> propsForDocs = new List<ObjectProperty[]>();
			List<OID> parents=new List<OID>();
			List<TPropertyTag> contTags=new List<TPropertyTag>();
			List<Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach[]> attaches = new List<Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach[]>();

			for (int i = 0; i < documentsForCreation.Length; i++)
			{
				BaseDocument itDoc=documentsForCreation[i];

				docIDs.Add(itDoc.ID);
				propsForDocs.Add(itDoc.PropertiesForCreation);
				parents.Add(folder);
				contTags.Add(Foss.FossDoc.ExternalModules.BusinessLogic.Schema.Folder.Attributes.Documents.Tag);
				attaches.Add(itDoc.Attaches);
			}

			session.ObjectDataManager.CreateObjectsWithOID(parents.ToArray(), contTags.ToArray(), docIDs.ToArray(), propsForDocs.ToArray());

			Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach.CreateAttaches(session, docIDs.ToArray(), attaches.ToArray());
			
			for (int i = 0; i < documentsForCreation.Length; i++)
			{
				BaseDocument itDoc = documentsForCreation[i];
				itDoc._AfterDocumentCreated();
			}

			return docIDs.ToArray();
		}

		/// <summary>
		/// Добавить набор свойств (вызвать перед созданием документа)
		/// </summary>
		/// <param name="objectProperty"></param>
		public void AddProperties(params ObjectProperty[] objectProperty)
		{
			foreach (var item in objectProperty)
			{
				_Props[item.PropertyTag] = item;
			}
		}

	}
}
