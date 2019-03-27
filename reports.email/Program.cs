using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Foss.FossDoc.ApplicationServer;
using DS;
using Converters = Foss.FossDoc.ApplicationServer.Converters;
using System.IO;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using Foss.FossDoc.ApplicationServer.Messaging;
using Foss.FossDoc.ExternalModules.BusinessLogic.Utils;
using Foss.FossDoc.ApplicationServer.IO;
using Foss.FossDoc.ApplicationServer.Converters.Restrictions;

namespace reports.email
{
	/// <summary>
	/// Програмна робота з звітами (про читанння, доставку) за поштовими повідомленнями. Цей приклад шукає звіт про доставку, а потім
	/// знаходить листа, якого цей звіт стосується.
	/// </summary>
	class Program
	{
		static void Main(string[] args)
		{
			try
			{

				string connectionString = "URL=corbaloc::1.2@andrewsalko:11301/AccessRoot;Login=Деловод;Password=123;AuthenticationAlgorithm=FossDoc;";

				using (ISession session = (ISession)Foss.FossDoc.ApplicationServer.Connection.Connector.Connect(connectionString))
				{
					var obj = session.ObjectDataManager;
					var msg = session.MessagingManager;

					//Цей приклад аналізує "Звіт про доставку", та знаходить листа, до якого цей звіт відноситься.
					//Якщо вам необхідно програмно відправити лист, див.приклад у проекті send.email.

					//Для тестування прикладу:
					//1) вам потрібні два користувача, з скриньками пошти (FossMail)
					//2) користувач Діловод надсилає листа (з опцією звіта про доставку та прочитання) - через клієнт FossDoc
					//3) інший користувач отримує листа, читає його
					//4) транспорт доставляє Діловоду звіти

					//5) цей приклад "читає" звіти з папки Вхідні Діловода, та знаходить у папці "Віправлені" лист-оригінал

					//Папка "Вхідні" поточного користувача - там будемо шукати звіти про доставку (нові, які ми ще не прочитали)
					OID inboxFolderID = msg.GetSpecialMessagingFolder(SpecialMessagingFolder.Inbox);

					//Папка "Відправлені" (там будемо шукати наші листи)
					OID sentFolderID = msg.GetSpecialMessagingFolder(SpecialMessagingFolder.SentMessages);

					//У звіта є своя "категорія" (тип об'єкту)
					//Звіт про доставку листа - Foss.FossDoc.ApplicationServer.Messaging.Schema.DeliveryReport.Category.OID
					//Звіт про не доставку листа -  Foss.FossDoc.ApplicationServer.Messaging.Schema.NondeliveryReport.Category.OID
					//Звіт про читання листа - Foss.FossDoc.ApplicationServer.Messaging.Schema.ReadReport.Category.OID
					
					//Класс PropertyRestrictionHelper дозволяє побудувати "фільтр"
					//Тег ObjectCategoryOID == DeliveryReport.Category.OID (перевірка "дорівнує" - relopEQ)
					PropertyRestrictionHelper propCategory = new PropertyRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID, Foss.FossDoc.ApplicationServer.Messaging.Schema.DeliveryReport.Category.OID, DS.relopEQ.ConstVal);
					TableRestrictionHelper tblCategory = new TableRestrictionHelper(DS.resProperty.ConstVal, propCategory);

					//також нам потрібно знайти не усі звіти, а лише ті, що ми "не читали". Властивість "Прочитано" (bool) може не бути встановлена, а також бути false.
					OID userID = session.AccessControlManager.RoleManager.ActAsUserOID;
					ExistRestrictionHelper ex = new ExistRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectReaded, userID);
					TableRestrictionHelper tlbEx = new TableRestrictionHelper(DS.resExist.ConstVal, ex);
					TableRestrictionHelper tblReadNotExists = new TableRestrictionHelper(DS.resNOT.ConstVal, tlbEx);

					//також класичне "ObjectReaded == false"
					PropertyRestrictionHelper propReadIsFalse = new PropertyRestrictionHelper(Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectReaded, false, DS.relopEQ.ConstVal);
					TableRestrictionHelper tblReadFalse = new TableRestrictionHelper(DS.resProperty.ConstVal, propReadIsFalse);

					//блок умов:
					TableRestrictionHelper tblOR = new TableRestrictionHelper(DS.resOR.ConstVal, tblReadFalse, tblReadNotExists);
					TableRestrictionHelper tblAND = new TableRestrictionHelper(DS.resAND.ConstVal, tblCategory, tblOR);

					OID[] foundReports = obj.GetChildren(new OID[] { inboxFolderID }, new TPropertyTag[] { Foss.FossDoc.ApplicationServer.Messaging.Schema.MessagesFolder.PR_CONTAINER_CONTENTS }, tblAND);
					if (foundReports != null && foundReports.Length == 0)
					{
						_Log("Нових звітів доставки не знайдено");
						return;
					}

					foreach (var reportID in foundReports)
					{
						_Log("Знайдено новий звіт:{0}", reportID.ToStringRepresentation());

						//У звіта отримуємо властивість PR_REPORT_TAG (бінарна властивість, зазвичай не дуже велика (наприклад 20 байт) де зберігається ідентифікатор повідомлення,
						//до якого йде звіт. Це як унікальний ключ, який можна використати для пошуку оригінального листа (який ми надіслали)
						//0x00310102

						var reportProps = obj.GetProperties(reportID, Foss.FossDoc.ApplicationServer.Messaging.Schema.Report.Attributes.PR_REPORT_TAG.Tag);
						if (reportProps != null && reportProps[0].PropertyTag.IsEquals(Foss.FossDoc.ApplicationServer.Messaging.Schema.Report.Attributes.PR_REPORT_TAG.Tag))
						{
							//тепер шукаємо у папці "Відправлені" наш лист
							//З технічного боку ми шукаємо "за бінарною властивістю" яка зараз у структурі ObjectProperty (reportProps[0].Value.GetbinVal())
							PropertyRestrictionHelper propReportTag = new PropertyRestrictionHelper(Foss.FossDoc.ApplicationServer.Messaging.Schema.Report.Attributes.PR_REPORT_TAG.Tag, reportProps[0], DS.relopEQ.ConstVal);
							TableRestrictionHelper tblPropReportTag = new TableRestrictionHelper(DS.resProperty.ConstVal, propReportTag);

							OID[] sentMessagesIDs = obj.GetChildren(new OID[] { sentFolderID }, new TPropertyTag[] { Foss.FossDoc.ApplicationServer.Messaging.Schema.MessagesFolder.PR_CONTAINER_CONTENTS }, tblPropReportTag);
							if (sentMessagesIDs != null && sentMessagesIDs.Length > 0)
							{
								//у нас тут буде лише 1 лист, який ми знайшли
								_Log("Знайдено оригінальний лист: {0}", sentMessagesIDs[0].ToStringRepresentation());

								//На цьому приклад завершено. Ви знайшли лист, який "було доставлено" успішно. Тепер ви можете зчитати з нього будь-які дані (Тема, файли..або інше)
							}
							else
							{
								_Log("Для цього звіта листа не знайдено. Можливо його було видалено з папки Відправлені.");
							}
						}

						//робимо його "прочитаним"
						msg.SetReadFlag(new OID[] { reportID }, true);
					}

					_Log("Завершення роботи");

				}//using
			}
			catch (Exception ex)
			{
				_Log(ex.ToString());
			}
		}

		static void _Log(string message, params object[] args)
		{
			Console.WriteLine(message, args);
		}

	}
}
