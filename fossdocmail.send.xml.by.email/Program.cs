using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using DS;
using Foss.FossDoc.ApplicationServer;
using Foss.FossDoc.ApplicationServer.Messaging;
using Foss.FossDoc.ApplicationServer.ObjectDataManagment;
using FossDocMail.ExternalModule;
using Converters = Foss.FossDoc.ApplicationServer.Converters;

namespace fossdocmail.send.xml.by.email
{
	/// <summary>
	/// Приклад демонструє створення листа (email) з XML документом (так званий оффлайн режим)
	/// </summary>
	/// <param name="args"></param>
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				//Увага! Цей приклад не є самостійним, вважається що ви вже вмієте створювати окремо сам документ, додавати файли-вкладення та ЕЦП.
				//Нижче розглядається випадок коли в базі документ вже існує, ви знаєте його ідентифікатор (OID)
				OID documentID = Converters.OID.FromString("000000008074ACBC7226467C90722EC6FF10BF96");

				string connectionString = "URL=corbaloc::1.2@adss:11301/AccessRoot;Login=Деловод;Password=123;AuthenticationAlgorithm=FossDoc;";

				using (ISession session = (ISession)Foss.FossDoc.ApplicationServer.Connection.Connector.Connect(connectionString))
				{
					var obj = session.ObjectDataManager;
					var msg = session.MessagingManager;

					if (!obj.GetExistance(documentID))
						throw new ApplicationException("Документ не знайдено в базі, створіть його та вкажіть свій ідентифікатор для цього тесту");

					//створити email (сам по собі він створюється у папці Вихідні, але посилання додаємо до документу - "Листування")
					OID emailID = _CreateEmail(obj, msg, "Передача електронного документу № X1");

					string tempDir = Foss.TemplateLibrary.IO.TempFolder.CreateTempFolder(Foss.TemplateLibrary.IO.TempFolder.FOSS_ON_LINE, "NBU_SND_XML");

					//менеджер IDocumentManager надається модулем FossDocMail (збірка FossDocMail.ExternalModule.Interfaces.dll)
					using (var docManager = (IDocumentManager)session.ExternalModulesManager.GetExternalModuleInterface(FossDocMail.ExternalModule.ExternalModule.Name, typeof(IDocumentManager).FullName))
					{
						//параметри "кому" (toOrgID, toDepartmentID) вкажіть ті, що вам потрібно (ці дані лише для тесту)

						string toOrgID = "F0000004"; //idHost
						string toDepartmentID = "f5e83799-4b82-47a5-b2b2-ac3191d5f808"; //idMailBox

						_CreateXMLDocument(obj, docManager, documentID, tempDir, emailID, toOrgID, toDepartmentID);

					}//docManager


					//додаємо посилання у документ - на створений email, його буде видно у вкладці "Листування"
					obj.AddChild(documentID, Foss.FossDoc.ApplicationServer.Messaging.Schema.ObjectWithMailing.Attributes.Mailing.Tag, emailID);

					//надсилаємо email:
					//msg.SubmitMessage(emailID);

				}//using session
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.ToString());
			}

		}//Main


		static OID _CreateEmail(IObjectDataManager obj, Foss.FossDoc.ApplicationServer.Messaging.IManager msg, string subject)
		{
			//Щоб надіслати лист, треба спочатку його створити (на сервері). Ми будемо створювати у папці "Вихідні".
			//отримаємо папку "Вихідні" поточного користувача:
			OID folderID = msg.GetSpecialMessagingFolder(SpecialMessagingFolder.Outbox);

			//propsForMessage - тут властивості нового листа, який буде створено у папці:
			List<ObjectProperty> propsForMessage = new List<ObjectProperty>
			{
				//категорія - "Вихідний лист" (email)
				Converters.Properties.ObjectPropertyBuilder.Create(Foss.FossDoc.ApplicationServer.Messaging.Schema.OutboundMessage.CategoryOID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID),
				//тема листа
				Converters.Properties.ObjectPropertyBuilder.Create(subject, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SUBJECT),
				//Важливість (0 - низька, 1- нормальна, 2 - висока)
				Converters.Properties.ObjectPropertyBuilder.Create(1, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_IMPORTANCE),
				//"IPM.Note" класс повідомлення
				Converters.Properties.ObjectPropertyBuilder.Create("IPM.Note", Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_MESSAGE_CLASS)
			};

			//кодування тіла
			Encoding enc = Encoding.UTF8;
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(enc.CodePage, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_INTERNET_CPID));

			string emailBody = "Короткий опис листа";
			byte[] bodyBytes = enc.GetBytes(emailBody);
			//тіло йде у бінарному вигляді
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(bodyBytes, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_BODY));

			//Додаємо властивості відправника:
			//Якщо у користувача кілька облікових записів пошти, беремо за замовчуванням:

			var defaultAddr = msg.GetCurrentUserDefaultAddressDescription();
			if (defaultAddr == null)
			{
				//беремо тільки адреси користувача, а не делегованих йому:
				var allAddrDescr = msg.GetCurrentUserAddressDescriptions(Foss.FossDoc.ApplicationServer.Messaging.UserAddresses.User);
				if (allAddrDescr != null)
				{
					//за допомогою linq знайдемо лише перший обл.запис де тип адреси "FMAIL" (наш приклад для пошти FossMail)
					//Якщо ви шукаєте звичайний email, то вкажіть SMTP.
					defaultAddr = allAddrDescr.FirstOrDefault(item => (item.Type != null && item.Type == "FMAIL"));
				}
			}

			if (defaultAddr == null)
			{
				throw new ApplicationException("Помилка, у користувача не знайдено облікових записів для роботи з поштою");
			}

			//властивості відправника:
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(defaultAddr.Address, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SENDER_EMAIL_ADDRESS));
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(defaultAddr.Type, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SENDER_ADDRTYPE));
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(defaultAddr.Name, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_SENDER_NAME));

			//якщо необхідно отримати звіти про доставку або прочитання листа:
			//Звіт про доставку потрібен
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(true, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED));
			//Звіт про прочитання потрібен
			propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(true, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_READ_RECEIPT_REQUESTED));

			OID emailMessageID = obj.CreateObject(folderID, Foss.FossDoc.ApplicationServer.Messaging.Schema.MessagesFolder.PR_CONTAINER_CONTENTS, propsForMessage.ToArray());

			return emailMessageID;
		}

		static void _CreateXMLDocument(IObjectDataManager obj, IDocumentManager mngr, OID documentID, string tempDirForXML, OID emailID, string toOrgID, string toDepartmentID)
		{
			var info = mngr.CreateXMLDocument(documentID, toOrgID, toDepartmentID);
			var msgID = info.XMLMsgID;

			string tempFolderToRemoveOnServer = info.TempFolderOnServer;//тут буде папка, яку потрібно передати серверу на видалення після закінчення роботи

			string xmlFile = string.Format("nbu.document.{0}.xml", msgID);
			string fullXMLFile = Path.Combine(tempDirForXML, xmlFile);

			using (var stream = Foss.FossDoc.ApplicationServer.IO.StreamAccessOptimizer.WrapStream(info.XMLStream))
			{
				if (stream == null)
					throw new ApplicationException("stream == null");

				using (FileStream fs = new FileStream(fullXMLFile, FileMode.Create))
				{
					stream.Position = 0;
					Foss.TemplateLibrary.StreamCopy.Copy(stream, fs);
				}
			}

			//створення файлу-вкладення
			var att = new Foss.FossDoc.ExternalModules.BusinessLogic.Utils.Attach(fullXMLFile);
			att.CreateAttach(obj, emailID, true);

			try
			{
				mngr.CleanupTemporaryData(tempFolderToRemoveOnServer);
			}
			catch (Exception)
			{
			}

			try
			{
				File.Delete(fullXMLFile);
			}
			catch
			{
			}

		}//_CreateXMLDocument


	}
}
