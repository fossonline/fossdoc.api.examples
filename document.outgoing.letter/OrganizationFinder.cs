using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DS;
using Foss.FossDoc.ApplicationServer;
using Foss.FossDoc.ApplicationServer.Converters.Restrictions;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using Conveters = Foss.FossDoc.ApplicationServer.Converters;

namespace document.outgoing.letter
{
	/// <summary>
	/// Пошук кореспондента за його властивістю
	/// </summary>
	class OrganizationFinder
	{
		IObjectDataManager _Obj;

		/// <summary>
		/// Тут ми шукаємо за іменем. Якщо ви хочете шукати за іншим полем, змінить тег, але пам'ятайте про типізацію
		/// </summary>
		static TPropertyTag _SearchTag = Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.DisplayName;

		/// <summary>
		/// Приклад: якщо вам потрібно шукати за кодом установи (НБУ) ідентифікатор організації - тег 0x3837001F використайте його замість _SearchTag
		/// </summary>
		static TPropertyTag _NBUHostIDTag = Conveters.PropertyTag.FromUInt32(0x3837001F);

		public OrganizationFinder(IObjectDataManager obj)
		{
			_Obj = obj ?? throw new ArgumentNullException(nameof(obj));
		}

		public OID FindOrganization(string searchKey)
		{
			OID result = Conveters.OID.Unspecified;	//якщо не знайдено, це буде результатом

			PropertyRestrictionHelper propOrgID = new PropertyRestrictionHelper(_NBUHostIDTag, searchKey, relopEQ.ConstVal);
			TableRestrictionHelper tbl = new TableRestrictionHelper(resProperty.ConstVal, propOrgID);

			OID[] korrIDs = _Obj.GetChildren(new OID[] { Foss.FossDoc.ExternalModules.EDMS.Schema.Dictionaries.KorrespondentsDictionary }, new TPropertyTag[] { Foss.FossDoc.ApplicationServer.Messaging.Schema.MessagesFolder.PR_CONTAINER_CONTENTS }, tbl);
			if (korrIDs != null && korrIDs.Length > 0)
			{
				result = korrIDs[0];	//Увага: можливо знайдеться декілька..тут ми беремо просто першого, але враховуйте це для вашої задачі
			}

			return result;
		}


	}
}
