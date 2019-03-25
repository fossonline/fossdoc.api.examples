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

namespace send.email
{
	/// <summary>
	/// Програмна відправка листа (email)
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
					//менеджер для роботи з об'єктами (через нього ви отримаєте будь-які властивості, наприклад Тема листа, відправник, наявність файлів та інше)
					var obj = session.ObjectDataManager;
					//менеджер для роботи з поштою
					var msg = session.MessagingManager;

					//Щоб надіслати лист, треба спочатку його створити (на сервері). Ми будемо створювати у папці "Вихідні".
					//отримаємо папку "Вихідні" поточного користувача:
					OID folderID = msg.GetSpecialMessagingFolder(SpecialMessagingFolder.Outbox);

					//тема
					string subject = "This is email subject";

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

					string emailBody = "Це зміст листа, просто текстовий. Hello this is email";
					byte[] bodyBytes = enc.GetBytes(emailBody);
					//тіло йде у бінарному вигляді
					propsForMessage.Add(Converters.Properties.ObjectPropertyBuilder.Create(bodyBytes, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_BODY));

					//Додаємо властивості відправника:
					//Якщо у користувача кілька облікових записів пошти, беремо за замовчуванням:

					var defaultAddr = msg.GetCurrentUserDefaultAddressDescription();
					if (defaultAddr == null)
					{
						//беремо тільки адреси користувача, а не делегованих йому:
						var allAddrDescr = msg.GetCurrentUserAddressDescriptions(UserAddresses.User);
						if (allAddrDescr != null)
						{
							//за допомогою linq знайдемо лише перший обл.запис де тип адреси "FMAIL" (наш приклад для пошти FossMail)
							//Якщо ви шукаєте звичайний email, то вкажіть SMTP.
							defaultAddr = allAddrDescr.FirstOrDefault(item => (item.Type != null && item.Type == "FMAIL"));
						}
					}

					if (defaultAddr == null)
					{
						_Log("Помилка, у користувача не знайдено облікових записів для роботи з поштою");
						return;
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

					//Тепер, коли обєкт листа вже створено на сервері, ми можемо додати файли (якщо це потрібно)
					//Для завантаження файлів використаємо утілітний класс з інтерфейсів бізнес-логіки.

					string fileName = @"C:\Documents\test.docx";

					//Варіант 1: це просто через утілітний класс, якщо не треба замислюватися як воно працює
					//Attach att = new Attach(fileName);
					//att.CreateAttach(session, emailMessageID);

					//Варіант 2: (дозволяє зрозуміти як саме програмно створюються файли у документах)
					_CreateAttach(obj, fileName, emailMessageID);

					//Лист створено, файл завантажено, тепер потрібно створити таблицю отримувачів:
					//У випадку використання транспорту FossMail адреса наприклад, може бути така:
					AddressDescription addr = new AddressDescription("C:UA/ADMD:CENTER_REGION/PRMD:CENTER/ORG:COMPANY/OU:OFFICE/PN:DELOVOD", "FMAIL", "Шевченко Тарас");

					//Властивості об'єкта "Отримувача":
					List<ObjectProperty> recipientProps = new List<ObjectProperty>
					{
						//категорія-відправник
						Converters.Properties.ObjectPropertyBuilder.Create(Foss.FossDoc.ApplicationServer.Messaging.Schema.Recipient.CategoryOID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID),
						//адреса
						Converters.Properties.ObjectPropertyBuilder.Create(addr.Address, Foss.FossDoc.ApplicationServer.Messaging.Schema.Address.PR_EMAIL_ADDRESS),
						//тип адреси
						Converters.Properties.ObjectPropertyBuilder.Create(addr.Type, Foss.FossDoc.ApplicationServer.Messaging.Schema.Address.PR_ADDR_TYPE),
						//імя 
						Converters.Properties.ObjectPropertyBuilder.Create(addr.Name, Foss.FossDoc.ApplicationServer.Messaging.Schema.Address.PR_DISPLAY_NAME),
						//тип (Кому)
						Converters.Properties.ObjectPropertyBuilder.Create(1, Foss.FossDoc.ApplicationServer.Messaging.Schema.Recipient.PR_RECIPIENT_TYPE),
						//тип обєкту
						Converters.Properties.ObjectPropertyBuilder.Create((int)ObjectType.MailUser, Foss.FossDoc.ApplicationServer.Messaging.Schema.Object.Attributes.PR_OBJECT_TYPE.Tag),
						//PR_SEARCH_KEY для отримувача
						Converters.Properties.ObjectPropertyBuilder.Create(Encoding.UTF8.GetBytes(string.Format("{0}:{1}", addr.Type.ToUpperInvariant(), addr.Address)), Foss.FossDoc.ApplicationServer.Messaging.Schema.Object.Attributes.PR_SEARCH_KEY.Tag)
					};

					//створили у повідомленні "Отримувача". Якщо треба, таким чином додайте й інших.
					obj.CreateObject(emailMessageID, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_MESSAGE_RECIPIENTS, recipientProps.ToArray());

					//А тепер вже "Відправка листа" - виклик SubmitMessage
					msg.SubmitMessage(emailMessageID);

					//щоб виглядало у "Відправлених" прочитаним:
					msg.SetReadFlag(new OID[] { emailMessageID }, true);

					_Log("Відправку листа виконано");

				}//using
			}
			catch (Exception ex)
			{
				_Log(ex.ToString());
			}
		}


		/// <summary>
		/// Функція демонструє як створити файл у поштовому повідомленні (або у документі). Також є класс Attach (в інтерфейсах бізнес-логіки) який дозволяє робити теж саме.
		/// </summary>
		/// <param name="obj"></param>
		/// <param name="fileName">Файл на диску</param>
		/// <param name="parentDocumentID">Документ або поштове повідомлення, де треба створити вкладення</param>
		static OID _CreateAttach(IObjectDataManager obj, string fileName, OID parentDocumentID)
		{
			string fileNameNonPath = Path.GetFileName(fileName);

			List<ObjectProperty> attachProps = new List<ObjectProperty>
			{
				//категория
				Converters.Properties.ObjectPropertyBuilder.Create(Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.CategoryOID, Foss.FossDoc.ApplicationServer.ObjectDataManagment.Schema.PropertyTags.ObjectCategoryOID),
				//имя файла
				Converters.Properties.ObjectPropertyBuilder.Create(fileNameNonPath, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_FILENAME),
				//длинное имя файла
				Converters.Properties.ObjectPropertyBuilder.Create(fileNameNonPath, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_LONG_FILENAME),
				//метод хранения
				Converters.Properties.ObjectPropertyBuilder.Create(1, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_METHOD),
				//метод - по значению (стандартный аттач)
				Converters.Properties.ObjectPropertyBuilder.Create(1, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_METHOD)
			};

			string ext = Path.GetExtension(fileName);
			if (!string.IsNullOrWhiteSpace(ext))
			{
				attachProps.Add(Converters.Properties.ObjectPropertyBuilder.Create(ext, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_EXTENSION));
			}

			OID attachID = obj.CreateObject(parentDocumentID, Foss.FossDoc.ApplicationServer.Messaging.Schema.Message.PR_MESSAGE_ATTACHMENTS, attachProps.ToArray());

			//тепер завантажуємо "Тіло" вкладення з файлу
			//Тут ми відкриємо та працюємо з файлом на диску, але ви можете робити це з будь-якого буферу
			using (StreamEx streamToWrite = obj.UpdateBinaryStream(attachID, Foss.FossDoc.ApplicationServer.Messaging.Schema.Attachment.PR_ATTACH_DATA_BIN, BinaryStreamAccessMode.CreateAndReadWrite))
			{
				using (FileStream fs = System.IO.File.OpenRead(fileName))
				{
					//буфер у нас 256  Кб
					Converters.StreamCopy.Perform(fs, streamToWrite, 1024 * 256);
				}
			}

			return attachID;
		}

		static void _Log(string message, params object[] args)
		{
			Console.WriteLine(message, args);
		}

	}
}
