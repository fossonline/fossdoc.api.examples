using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Foss.FossDoc.ExternalModules.BusinessLogic.Registration;
using Foss.TemplateLibrary;
using DS;
using Converters = Foss.FossDoc.ApplicationServer.Converters;
using Foss.FossDoc.ApplicationServer;
using Foss.FossDoc.ExternalModules.EDMS.Interfaces;
using System.Diagnostics;

namespace document.outgoing.letter
{
	public class OutgoingLetterWithAutoRegistration : BaseDocument
	{
		/// <summary>
		/// Зміст документа
		/// </summary>
		string _Content;

		/// <summary>
		/// Дата реєстрації документа
		/// </summary>
		DateTime _DocumentDate;

		/// <summary>
		/// Кореспонденти (або один, мінімум)
		/// </summary>
		OID[] _Correspondents;


		public OutgoingLetterWithAutoRegistration(ISession session, OID parentFolderID, DateTime docDate, string content, OID[] correspondents) : base(session, parentFolderID, Foss.FossDoc.ExternalModules.EDMS.Schema.CA_OutgoingLetter.CategoryId)
		{
			if (string.IsNullOrWhiteSpace(content))
				throw new ArgumentException(nameof(content));

			if(docDate==DateTime.MinValue)
				throw new ArgumentException(nameof(docDate));

			if(correspondents==null || correspondents.Length==0)
				throw new ArgumentException(nameof(correspondents));

			_Content = content;
			_DocumentDate = docDate;
			_Correspondents = correspondents;
		}

		protected override void _BeforeDocumentCreated()
		{
			base._BeforeDocumentCreated();

			_Props[Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_OutgoingLetter.EDMS_FIELD_DATE_CREATION] = Converters.Properties.ObjectPropertyBuilder.Create(_DocumentDate, Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_OutgoingLetter.EDMS_FIELD_DATE_CREATION);
			_Props[Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_OutgoingLetter.EDMS_FIELD_DOCUMENT_CONTENT] = Converters.Properties.ObjectPropertyBuilder.Create(_Content, Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_OutgoingLetter.EDMS_FIELD_DOCUMENT_CONTENT);
			_Props[Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_OutgoingLetter.EDMS_FIELD_KORRESPONDENTS] = Converters.Properties.ObjectPropertyBuilder.Create(_Correspondents, Foss.FossDoc.ExternalModules.EDMS.Schema.PropertyTags.CA_OutgoingLetter.EDMS_FIELD_KORRESPONDENTS);

		}

		protected override void _AfterDocumentCreated()
		{
			base._AfterDocumentCreated();

			//Наш документ буде зареєстровано негайно після створення в базі. У вас повинен існувати нумератор (у підрозділі)
			using (var regMgr = (Foss.FossDoc.ExternalModules.BusinessLogic.Registration.IRegistrationManager2)_Session.ExternalModulesManager.GetExternalModuleInterface(Foss.FossDoc.ExternalModules.BusinessLogic.Schema.ExternalModule.Name, typeof(Foss.FossDoc.ExternalModules.BusinessLogic.Registration.IRegistrationManager2).FullName))
			{
				regMgr.RegisterDocument(ID);
			}
		}
	}
}