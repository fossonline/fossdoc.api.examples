using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using DS;
using Foss.FossDoc.ApplicationServer;
using Foss.FossDoc.ApplicationServer.Converters.Restrictions;
using Conveters = Foss.FossDoc.ApplicationServer.Converters;
using System.IO;

namespace document.outgoing.letter
{
	/// <summary>
	/// Створює об'єкт цифрового підпису з вашого файла (на диску)
	/// </summary>
	class DigitalSignature
	{
		IObjectDataManager _Obj;

		const string _BODY_ATTRIB_NAME = "Body";

		public DigitalSignature(IObjectDataManager obj)
		{
			_Obj = obj ?? throw new ArgumentNullException(nameof(obj));
		}

		public OID CreateSignatureRecord(OID parentFileID, string fileName, OID userAuthorID, string digitalSignFile)
		{
			//Увага! ЕЦП файл зазвичай невеликий - 3-4 Кб (відокремлена ЕЦП), тому ми можемо його встановити як звичайну властивість. Для файлів-документів так робити не можна.
			byte[] signBody = File.ReadAllBytes(digitalSignFile);

			string signFileName = Path.GetFileName(digitalSignFile);

			string[] attribs = new string[] { _BODY_ATTRIB_NAME };

			//перелік властивостей для запису ЕЦП (для файлу)
			ObjectProperty[] propsForCreation =
			{
				Conveters.Properties.ObjectPropertyBuilder.Create(fileName, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName),
				Conveters.Properties.ObjectPropertyBuilder.Create(Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Category.OID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID),

				//Тип ЕЦП - Файл (константа)
				Conveters.Properties.ObjectPropertyBuilder.Create(Foss.FossDoc.ExternalModules.BusinessLogic.DigitalSignature.Schema.DSType.Objects.AttachedFileID, Foss.FossDoc.ExternalModules.BusinessLogic.DigitalSignature.Schema.ObjectDigitalSignature.Attributes.DSType.Tag),

				//Автор ЕЦП (умовно - будь-який користувач)
				Conveters.Properties.ObjectPropertyBuilder.Create(userAuthorID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Attributes.SignatureAuthor.Tag),

				//Тіло ЕЦП
				Conveters.Properties.ObjectPropertyBuilder.Create(signBody, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Attributes.Signature.Tag),

				//відокремлений файл цифрового підпису (зазвичай у вас це .p7s)
				Conveters.Properties.ObjectPropertyBuilder.Create(signFileName, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Attributes.SeparateSignatureFile.Tag),

				//посилання на об'єкт-файл який було підписано
				Conveters.Properties.ObjectPropertyBuilder.Create(parentFileID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Attributes.SignedObject.Tag),

				//константа - Body для атрибутів, які було підписано
				Conveters.Properties.ObjectPropertyBuilder.Create(attribs, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Attributes.SignedObjectAttributes.Tag),
				Conveters.Properties.ObjectPropertyBuilder.Create(attribs, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.DigitalSignature.Attributes.SignedObjectAttributesDisplayNames.Tag)
			};

			OID id = _Obj.CreateObject(parentFileID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.DigitalSignature.Schema.ObjectWithDigitalSignatures.Attributes.DigitalSignatures.Tag, propsForCreation);

			return id;
		}

	}
}
