using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DS;
using Foss.FossDoc.ApplicationServer;
using Foss.FossDoc.ApplicationServer.Converters.Restrictions;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using Converters = Foss.FossDoc.ApplicationServer.Converters;

namespace document.incoming.processing
{
	class IncomingAnalyser
	{
		/// <summary>
		/// Поле (тип short) де буде значення 1 яке позначає що цей документ ми вже обробляли
		/// (приховане поле)
		/// </summary>
		static readonly TPropertyTag _PROCESSED_MARK_TAG = Converters.PropertyTag.FromUInt32(0x40150002);

		const short _PROCESSED_MARK_VALUE = 1;

		const int _BLOCK_SIZE = 30;

		ISession _Session;
		OID _IncomingLettersFolderID;
		string _FolderForDocuments;

		public IncomingAnalyser(ISession session, OID incomingLettersFolderID, string destFolderForDocuments)
		{
			_Session = session ?? throw new ArgumentNullException(nameof(session));
			_IncomingLettersFolderID = incomingLettersFolderID;
			_FolderForDocuments = destFolderForDocuments;
		}

		public int Perform()
		{
			//Ідея обробки: знайти всі документи, які ми не переглядали та виконати з ними необхідні дії.
			//далі, для кожного проставити приховану властивість (помітку), щоб не знайти його наступного разу.

			var obj = _Session.ObjectDataManager;

			ExistRestrictionHelper flagExists = new ExistRestrictionHelper(_PROCESSED_MARK_TAG);
			TableRestrictionHelper tblExists = new TableRestrictionHelper(DS.resExist.ConstVal, flagExists);
			TableRestrictionHelper tblNot = new TableRestrictionHelper(DS.resNOT.ConstVal, tblExists);

			OID[] docIDs = obj.GetChildren(_IncomingLettersFolderID, Foss.FossDoc.ExternalModules.BusinessLogic.Schema.Folder.Attributes.Documents.Tag, tblNot);
			if (docIDs == null || docIDs.Length == 0)
				return 0; //нових документів не знайдено

			int processedCount = 0;

			//Щоб не навантажити сервер (якщо документів виявиться дуже багато), розбиваємо на блоки та по-блочно опрацюємо
			List<OID> processList = new List<OID>();
			for (int i = 0; i < docIDs.Length; i++)
			{
				processList.Add(docIDs[i]);
				if (processList.Count == _BLOCK_SIZE)
				{
					_Process(processList.ToArray());
					processedCount += processList.Count;
					processList.Clear();
				}
			}

			if (processList.Count > 0)
			{
				_Process(processList.ToArray());
				processedCount += processList.Count;
			}

			return processedCount;
		}

		static readonly TPropertyTag[] _DOC_TAGS =
		{
			Foss.FossDoc.ApplicationServer.Messaging.Schema.ObjectWithAttachedFile.Attributes.AttachedFiles.Tag,	//0 файли-вкладення
			Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_DOCUMENT_CONTENT,	//1 Зміст документу
			Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_KORRESPONDENT,		//2 Кореспондент
			Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.PR_REGNUMBER_SOURCE,			//3 Індекс документу (чужий реєстр номер)
			Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_CREATION_DATE_SOURCE,//4 Дата документу (чужа реєстр.дата)
			Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.PR_REGNUMBER,					//5 Реєс.номер (Індекс надходження)
			Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_DATE_CREATION		//6 Дата надходження (дата реєстрації)
		};

		static readonly TPropertyTag _IdHostTag = Converters.PropertyTag.FromUInt32(0x3837001F);

		static readonly TPropertyTag[] _CORR_TAGS =
		{
			Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName,	//0 Ім'я кореспондента (назва)
			_IdHostTag																			//1 IdHost код ППР
		};

		void _Process(OID[] docIDs)
		{
			var obj = _Session.ObjectDataManager;

			var allProps = obj.GetProperties(docIDs, _DOC_TAGS);

			for (int i = 0; i < docIDs.Length; i++)
			{
				var props = allProps[i];
				if (props == null)
					continue;

				OID[] attachIDs = null;
				if (props[0].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.Messaging.Schema.ObjectWithAttachedFile.Attributes.AttachedFiles.Tag))
				{
					attachIDs = props[0].Value.GetMVoidVal();
				}

				string content = string.Empty;
				if (props[1].PropertyTag.IsEquals(Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_DOCUMENT_CONTENT))
				{
					content = props[1].Value.GetstrVal();
				}

				OID corrID = Converters.OID.Unspecified;
				if (props[2].PropertyTag.IsEquals(Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_KORRESPONDENT))
				{
					corrID = props[2].Value.GetoidVal();
				}

				string docIndex = string.Empty;
				if (props[3].PropertyTag.IsEquals(Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.PR_REGNUMBER_SOURCE))
				{
					docIndex = props[3].Value.GetstrVal();
				}

				DateTime docDate = DateTime.MinValue;
				if (props[4].PropertyTag.IsEquals(Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_CREATION_DATE_SOURCE))
				{
					docDate = DateTime.FromFileTime(props[4].Value.GettVal()).ToLocalTime();
				}

				string docRegNumber = string.Empty;
				if (props[5].PropertyTag.IsEquals(Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.PR_REGNUMBER))
				{
					docRegNumber = props[5].Value.GetstrVal();
				}

				DateTime docRegDate = DateTime.MinValue;
				if (props[6].PropertyTag.IsEquals(Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_IncomingLetter.EDMS_FIELD_DATE_CREATION))
				{
					docRegDate = DateTime.FromFileTime(props[6].Value.GettVal()).ToLocalTime();
				}

				Console.WriteLine($"Document: {docIDs[i].ToStringRepresentation()} {content}  corrID:{corrID.ToStringRepresentation()}  {docIndex} ({docDate}) -- registered {docRegNumber} ({docRegDate})");

				string corrName = corrID.ToStringRepresentation();
				string idHost = string.Empty;

				var corrProps = obj.GetProperties(corrID, _CORR_TAGS);
				if (corrProps != null)
				{
					if (corrProps[0].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName))
					{
						corrName = corrProps[0].Value.GetstrVal();
					}

					if (corrProps[1].PropertyTag.IsEquals(_IdHostTag))
					{
						idHost = corrProps[1].Value.GetwstrVal();   //особливий випадок властивості Getwstr якщо тег ..001F
					}
				}


				//Тут ви можете оперувати з даними документа та робити будь-що (зберігати на диск, викликати через API інші системи)
				//Для прикладу зберігаємо на диск в окремі папки кожен документ та його файли
				_SaveDocumentInfo(obj, docIDs[i], corrID, corrName, idHost, attachIDs, content, docIndex, docDate, docRegNumber, docRegDate);

			}//for

			//Вважаємо, що блок документів оброблено - та помічаємо їх "флагом" щоб в наступному разі їх не знайти
			_MarkDocumentsAsProcessed(docIDs);

		}//_Process

		void _MarkDocumentsAsProcessed(OID[] docIDs)
		{
			ObjectProperty prop = Converters.Properties.ObjectPropertyBuilder.Create(_PROCESSED_MARK_VALUE, _PROCESSED_MARK_TAG);

			using (var session = _Session.Clone())
			{
				var obj = session.ObjectDataManager;

				using (var disp = obj.BeginIgnoreObjectLocks())
				{
					obj.SetProperties(docIDs, prop);
				}
			}
		}

		void _SaveDocumentInfo(IObjectDataManager obj, OID docID, OID corrID, string corrName, string idHost, OID[] attachIDs, string content, string docIndex, DateTime docDate, string docRegNumber, DateTime docRegDate)
		{
			string dirName = docID.ToStringRepresentation();    //папка буде як ідентифікатор документу це надасть унікальності

			string destDir = Path.Combine(_FolderForDocuments, dirName);
			Directory.CreateDirectory(destDir);

			string fileNameInfo = $"document_info_{docID.ToStringRepresentation()}.txt";
			string fullFileName = Path.Combine(destDir, fileNameInfo);

			using (var fs = File.CreateText(fullFileName))
			{
				fs.WriteLine($"ID: {docID.ToStringRepresentation()}");

				fs.WriteLine($"Кореспондент: {corrName} ({corrID.ToStringRepresentation()})");
				fs.WriteLine($"Кореспондент IdHost: {idHost}");

				fs.WriteLine($"Content: {content}");
				fs.WriteLine($"Індекс документу: {docIndex}");
				fs.WriteLine($"Дата документу: {docDate}");
				fs.WriteLine($"Індекс надходження: {docRegNumber}");	//реєстр.номер в вашій установі
				fs.WriteLine($"Дата надходження: {docRegDate}");
			}

			//тепер збережемо на диск всі файли-вкладення які є в картці
			//Тут нам допоможе утілітний готовий класс:

			if (attachIDs != null)
			{
				for (int i = 0; i < attachIDs.Length; i++)
				{
					//Скачає файл на диск (ім'я файлу буде таке як в базі, 0 означає що розмір блоку вам не важливий - буде за замовчанням 256 Кб)
					Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach.DownloadAttach(obj, attachIDs[i], destDir, null, 0);
				}
			}

		}

	}
}
