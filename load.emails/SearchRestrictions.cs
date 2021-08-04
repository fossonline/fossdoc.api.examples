using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Foss.FossDoc.ApplicationServer.Converters.Restrictions;
using DS;
using Converters = Foss.FossDoc.ApplicationServer.Converters;

namespace load.emails
{
	/// <summary>
	/// Демонструє різні варіанти пошукових фільтрів (restrict), які можна використати для читання документів або поштових листів
	/// </summary>
	class SearchRestrictions
	{
		/// <summary>
		/// Пошук непрочитаних листів
		/// </summary>
		/// <returns></returns>
		public static DS.TableRestriction GetRestrictForNonReadMessages()
		{
			//Нижче ми готуємо умову для пошуку (restrict).
			//Умова "Прочитано" НЕ рівно true. Але це спрацює лише коли властивість встановлено (а для нового листа його якраз не буде)
			PropertyRestrictionHelper propRead = new PropertyRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectReaded, true, DS.relopNE.ConstVal);
			TableRestrictionHelper tblNotReadLetters = new TableRestrictionHelper(DS.resProperty.ConstVal, propRead);

			//а ось цей блок відбирає "властивість Прочитано не знайдено"
			ExistRestrictionHelper readExists = new ExistRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectReaded);
			TableRestrictionHelper tblReadExists = new TableRestrictionHelper(DS.resExist.ConstVal, readExists);
			TableRestrictionHelper tblNot = new TableRestrictionHelper(DS.resNOT.ConstVal, tblReadExists);

			//Комплексна умова OR містить дві умови: "Прочитано НЕ є true" та "Прочитано не існує" 
			TableRestrictionHelper tblMain = new TableRestrictionHelper(DS.resOR.ConstVal, tblNot, tblNotReadLetters);

			return tblMain;
		}

		/// <summary>
		/// Діапазон дати (від...до). Увага: дата-час містить хвилини та секунди. У разі, якщо вам потрібно шукати "за датою",
		/// рекомендується брати від 00:00:00 до 23:59:59 кінцевої дати, таким чином усі необхідні документи потраплять до виборки.
		/// </summary>
		/// <param name="start">Дата-час початку</param>
		/// <param name="end">Дата-час закінчення</param>
		/// <param name="searchTag">Поле</param>
		/// <returns></returns>
		public static DS.TableRestriction GetDateRangeRestriction(DateTime start, DateTime end, TPropertyTag searchTag)
		{
			var resStart = GetRestrictionStartDate(start, searchTag);
			var resEnd = GetRestrictionEndDate(end, searchTag);

			//обєднуємо "AND" умовою
			TableRestrictionHelper tblAND = new TableRestrictionHelper(DS.resAND.ConstVal, resStart, resEnd);
			return tblAND;
		}

		/// <summary>
		/// Пошук листів у яких поле searchTag >= fromDate
		/// </summary>
		/// <param name="fromDate"></param>
		/// <param name="searchTag"></param>
		/// <returns></returns>
		public static IRestrictionHelper GetRestrictionStartDate(DateTime fromDate, TPropertyTag searchTag)
		{
			//relopGE - "більше або дорівнює"
			//relopGT - "більше"

			PropertyRestrictionHelper startDateRestriction = new PropertyRestrictionHelper(searchTag, fromDate, DS.relopGE.ConstVal);
			return startDateRestriction;
		}

		/// <summary>
		/// Пошук листів, у яких поле searchTag <= toDate
		/// </summary>
		/// <param name="toDate"></param>
		/// <param name="searchTag"></param>
		/// <returns></returns>
		public static IRestrictionHelper GetRestrictionEndDate(DateTime toDate, TPropertyTag searchTag)
		{
			//relopLE - "менше або дорівнює"
			//relopLT - "менше"

			PropertyRestrictionHelper endDateRestriction = new PropertyRestrictionHelper(searchTag, toDate, DS.relopLE.ConstVal);
			return endDateRestriction;
		}


	}
}
