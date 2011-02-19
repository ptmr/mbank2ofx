package pl.sevencoins.mbanktools;
/**
 * mbannCSV converter
 * License: GPL v3.0
 * http://www.gnu.org/licenses/gpl-3.0.txt
 * 
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.SortedSet;
import java.util.TimeZone;
import java.util.TreeSet;
import java.util.UUID;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JFileChooser;

import joptsimple.OptionParser;
import joptsimple.OptionSet;
import net.sf.ofx4j.domain.data.ResponseEnvelope;
import net.sf.ofx4j.domain.data.ResponseMessageSet;
import net.sf.ofx4j.domain.data.banking.AccountType;
import net.sf.ofx4j.domain.data.banking.BankAccountDetails;
import net.sf.ofx4j.domain.data.banking.BankStatementResponse;
import net.sf.ofx4j.domain.data.banking.BankStatementResponseTransaction;
import net.sf.ofx4j.domain.data.banking.BankingResponseMessageSet;
import net.sf.ofx4j.domain.data.common.Status;
import net.sf.ofx4j.domain.data.common.Status.KnownCode;
import net.sf.ofx4j.domain.data.common.Transaction;
import net.sf.ofx4j.domain.data.common.TransactionList;
import net.sf.ofx4j.domain.data.common.TransactionType;
import net.sf.ofx4j.domain.data.signon.SignonResponse;
import net.sf.ofx4j.domain.data.signon.SignonResponseMessageSet;
import net.sf.ofx4j.io.AggregateMarshaller;
import net.sf.ofx4j.io.OFXWriter;
import net.sf.ofx4j.io.v2.OFXV2Writer;

import com.csvreader.CsvReader;

public class MBankCSVToOFX {

    private static final String TOD_OFFSET = " 12:00"; // date conversion works
						       // better if transactions
						       // are made in the middle
						       // of the day
    private static final String INPUT_FILES_CHARSET = "windows-1250";
    private static Map<String, TransactionType> transactionTypes = new HashMap<String, TransactionType>();
    static final Logger log = Logger.getLogger(MBankCSVToOFX.class
	    .getCanonicalName());

    static String hexEncode(byte[] aInput) {
	final StringBuilder result = new StringBuilder();
	final char[] digits = { '0', '1', '2', '3', '4', '5', '6', '7', '8',
		'9', 'a', 'b', 'c', 'd', 'e', 'f' };
	for (int idx = 0; idx < aInput.length; ++idx) {
	    final byte b = aInput[idx];
	    result.append(digits[(b & 0xf0) >> 4]);
	    result.append(digits[b & 0x0f]);
	}
	return result.toString();
    }

    /**
     * @param args
     */
    public static void main(String[] args) {

	File inputFile = null;
	FileInputStream fis = null;
	String inputFileName = null;
	String outputFileName = null;
	OptionParser optionParser = new OptionParser("nf:ho:");
	OptionSet options = optionParser.parse(args);
	if (options.has("h")) {
	    System.out
		    .println("Usage: MBankCSVToOFX -f inputFilename -o outputFilename -n (append note to name) -h (help)");
	    System.exit(0);
	}
	boolean appendNoteToName = options.has("n");
	inputFileName = (String) options.valueOf("f");
	outputFileName = (String) options.valueOf("o");
	initMap();
	try {
	    final Logger log = Logger.getLogger(MBankCSVToOFX.class
		    .getCanonicalName());
	    if (inputFileName == null) {
		final JFileChooser chooser = new JFileChooser();
		chooser.setDialogTitle("Please chooose mBank CSV file");
		final int returnVal = chooser.showOpenDialog(null);
		if (returnVal == JFileChooser.APPROVE_OPTION) {
		    inputFile = chooser.getSelectedFile();
		    if (inputFile != null)
			log.info("You chose to open this file: "
				+ chooser.getSelectedFile().getName());
		    else {
			log.warning("You have not chosen a file to open, exiting");
			System.exit(0);
		    }
		} else {
		    log.warning("You have not chosen a file to open, exiting");
		    System.exit(0);
		}

	    } else
		inputFile = new File(inputFileName);
	    log.setLevel(Level.ALL);
	    fis = new FileInputStream(inputFile);

	    final Hashtable<String, String> global = new Hashtable<String, String>();

	    final MessageDigest sha = MessageDigest.getInstance("SHA-1");

	    log.info("Opening file:" + inputFile.getName());

	    fis = new FileInputStream(inputFile);
	    final CsvReader reader = new CsvReader(fis,
		    Charset.forName(INPUT_FILES_CHARSET));
	    final AggregateMarshaller marshaller = new AggregateMarshaller();
	    log.info("Reading records");

	    final List<String[]> transactions = readRecords(global, reader);
	    if (transactions.size() == 0) {
		log.warning("No transactions, possible bad file format");
		System.exit(1);

	    } else {
		log.info("Processing " + transactions.size() + " transactions");
	    }
	    if (outputFileName == null)
		outputFileName = inputFile.getAbsolutePath() + ".ofx";
	    log.info("Opening output file:" + outputFileName);

	    final OFXWriter ofxWriter = new OFXV2Writer(new FileWriter(
		    outputFileName));
	    ((OFXV2Writer) ofxWriter).setWriteAttributesOnNewLine(true);
	    log.info("Generating ofx structures");

	    final SortedSet<ResponseMessageSet> responseMsgs = new TreeSet<ResponseMessageSet>();
	    final SignonResponseMessageSet srms = new SignonResponseMessageSet();
	    final SignonResponse signonResponse = new SignonResponse();
	    signonResponse.setLanguage("Polish");
	    final Status statusOk = new Status();
	    statusOk.setCode(KnownCode.SUCCESS);
	    signonResponse.setStatus(statusOk);
	    signonResponse.setTimestamp(new Date());
	    srms.setSignonResponse(signonResponse);
	    responseMsgs.add(srms);

	    final List<Transaction> tList = new ArrayList<Transaction>();
	    DateFormat formatter = new SimpleDateFormat("yyyy-MM-dd H:m");
	    formatter.setTimeZone(TimeZone.getTimeZone("CET"));
	    for (final String[] i : transactions) {
		final String block_date = i[0];
		final String accounting_date = i[1];
		String operationDescription = i[2];
		String payeeData = i[3];
		String payeeAccount = i[4];
		String operationNote = i[5];
		Double charge = trimMoneyValue(i[6], global.get("currency"));

		final Transaction transaction = new Transaction();
		transaction.setAmount(charge);
		transaction.setDateInitiated(formatter.parse(block_date
			+ TOD_OFFSET));
		transaction.setDatePosted(formatter.parse(block_date
			+ TOD_OFFSET));
		transaction.setDateAvailable(formatter.parse(accounting_date
			+ TOD_OFFSET));

		payeeAccount = payeeAccount.replaceAll("'", "");
		operationNote = operationNote.replaceAll("\"", "");

		operationDescription = removeOgonki(operationDescription);
		payeeData = removeOgonki(payeeData);
		operationNote = removeOgonki(operationNote);
		int operNo = 0;
		boolean operationTypeFound = false;
		for (final String oper : transactionTypes.keySet()) {
		    if (operationDescription.startsWith(oper)) {
			transaction.setTransactionType(transactionTypes
				.get(oper));
			// description = description.substring(oper.length());
			operationTypeFound = true;
			transaction.setName(operationDescription);
			break;
		    }
		    operNo++;
		}

		transaction.setPayeeId(payeeAccount);

		if (payeeAccount.length() > 10) {
		    final BankAccountDetails bankAccountTo = new BankAccountDetails();
		    bankAccountTo.setAccountNumber(payeeAccount);
		    if (payeeAccount.length() == 26)
			bankAccountTo.setBankId(payeeAccount.substring(2, 10));
		    else
			bankAccountTo.setBankId(payeeAccount.substring(0, 8));

		    bankAccountTo.setAccountType(AccountType.CHECKING);
		    transaction.setBankAccountTo(bankAccountTo);
		}
		if (!operationTypeFound) {
		    transaction.setTransactionType(TransactionType.DEBIT);
		    log.info("Unknown operation type:" + operationDescription
			    + ", using defaults");
		    transaction.setName(operationDescription);
		}
		if (appendNoteToName) {
		    transaction.setName(transaction.getName() + ":"
			    + operationNote.trim());
		}

		transaction.setMemo(removeMultipleSpaces(operationNote.trim()));
		final String toSign = i[0] + i[1] + i[2] + i[3] + i[4] + i[5]
			+ i[6];
		transaction.setId(hexEncode(sha.digest(toSign.getBytes())));
		tList.add(transaction);
	    }

	    final TransactionList transactionList = new TransactionList();
	    formatter = new SimpleDateFormat("dd.MM.yy");
	    if (global.get("start_date") != null)
		transactionList.setStart(formatter.parse(global
			.get("start_date")));
	    if (global.get("end_date") != null)
		transactionList.setEnd(formatter.parse(global.get("end_date")));
	    transactionList.setTransactions(tList);
	    final BankAccountDetails bac = new BankAccountDetails();
	    bac.setAccountNumber(global.get("account_number").replaceAll(" ",
		    ""));
	    bac.setAccountType(AccountType.CHECKING);
	    bac.setBankId("BREXPLPWMBK");

	    final BankStatementResponse statementResponse = new BankStatementResponse();
	    statementResponse.setAccount(bac);
	    statementResponse.setCurrencyCode(global.get("currency"));
	    statementResponse.setTransactionList(transactionList);

	    final BankStatementResponseTransaction bsrt = new BankStatementResponseTransaction();
	    bsrt.setMessage(statementResponse);
	    bsrt.setStatus(statusOk);
	    bsrt.setUID(UUID.randomUUID().toString());

	    final BankingResponseMessageSet brms = new BankingResponseMessageSet();

	    brms.setStatementResponse(bsrt);
	    responseMsgs.add(brms);

	    final ResponseEnvelope response = new ResponseEnvelope();
	    response.setMessageSets(responseMsgs);
	    log.info("Writing ofx file");

	    marshaller.marshal(response, ofxWriter);
	    ofxWriter.close();

	} catch (final FileNotFoundException e) {
	    log.log(Level.SEVERE, "Error while converting", e);
	} catch (final IOException e) {
	    log.log(Level.SEVERE, "Error while converting", e);
	} catch (final ParseException e) {
	    log.log(Level.SEVERE, "Error while converting", e);
	} catch (final NoSuchAlgorithmException e) {
	    log.log(Level.SEVERE, "Error while converting", e);
	}

    }

    private static void initMap() {
	transactionTypes.put("ZERWANIE LOKATY TERMINOWEJ", TransactionType.DEP);
	transactionTypes.put("ODSETKI", TransactionType.INT);
	transactionTypes.put("OPLATA", TransactionType.SRVCHG);
	transactionTypes.put("PODATEK", TransactionType.DEBIT);
	transactionTypes.put("KAPITALIZACJA", TransactionType.INT);
	transactionTypes.put("ABONAMENT", TransactionType.REPEATPMT);
	transactionTypes.put("PROWIZJE", TransactionType.SRVCHG);
	transactionTypes.put("PROWIZJA", TransactionType.SRVCHG);
	transactionTypes.put("PRZELEW", TransactionType.XFER);
	transactionTypes.put("WYPLATA", TransactionType.CASH);
	transactionTypes.put("ZAKUP", TransactionType.DEBIT);
	transactionTypes.put("OPL. ZA ZLECENIE STALE", TransactionType.SRVCHG);

	transactionTypes.put("KREDYT - SPLATA RATY", TransactionType.DEBIT);
	transactionTypes.put("KREDYT-SKLADKA ZA UBEZPIECZENIE",
		TransactionType.DEBIT);
	transactionTypes.put("POS ZWROT TOWARU", TransactionType.CREDIT);
	transactionTypes.put("SKLADKA", TransactionType.REPEATPMT);
	transactionTypes.put("RECZNA SPLATA KARTY KREDYT",
		TransactionType.DEBIT);
	transactionTypes
		.put("AUTOMATYCZNA SPLATA KARTY", TransactionType.DEBIT);
	transactionTypes.put("STORNO OBCIAZENIA", TransactionType.CREDIT);
	transactionTypes.put("KREDYT - WCZESNIEJSZA SPLATA",
		TransactionType.CREDIT);
	transactionTypes.put("WPLATA WE WPLATOMACIE", TransactionType.CREDIT);
	transactionTypes.put("KREDYT - UZNANIE", TransactionType.CREDIT);
	transactionTypes.put("KREDYT - PROW. URUCHOMIENIE",
		TransactionType.SRVCHG);

    }

    static NumberFormat nf = NumberFormat.getInstance(new Locale("pl", "PL"));

    private static Double trimMoneyValue(String input, String currency) {

	try {
	    return nf.parse(input.replaceAll(" ", "")).doubleValue();
	} catch (ParseException e) {
	    return 0.0;
	}
	// return input.replaceAll(" ", "").replaceAll(",", ".")
	// .replaceAll(currency, "");
    }

    private static List<String[]> readRecords(
	    final Hashtable<String, String> global, final CsvReader reader)
	    throws IOException {
	final List<String[]> transactions = new ArrayList<String[]>();
	reader.setDelimiter(';');
	reader.setUseComments(false);
	reader.setTextQualifier('\'');

	int state = 0;

	while (reader.readRecord()) {
	    final String[] record = reader.getValues();
	    final String first = record[0];
	    switch (state) {
	    case 0:
		if (first.startsWith("#Klient")) {
		    if (reader.readRecord()) {
			global.put("name", reader.getValues()[0]);
			state = 1;
		    }
		}
		break;
	    case 1:
		if (first.startsWith("#Za okres:")) {
		    if (reader.readRecord()) {
			global.put("start_date", reader.getValues()[0]);
			global.put("end_date", reader.getValues()[1]);
			state = 2;
		    }
		}
		break;
	    case 2:
		if (first.startsWith("#Waluta")) {
		    if (reader.readRecord()) {
			global.put("currency", reader.getValues()[0]);
			state = 3;
		    }
		}
		break;
	    case 3:
		if (first.startsWith("#Numer rachunku")) {
		    if (reader.readRecord()) {
			global.put("account_number", reader.getValues()[0]);
			state = 4;
		    }
		}
		break;
	    case 4:
		if (first.startsWith("#Saldo pocz")) {
		    global.put("start_balance", reader.getValues()[1]);
		    state = 5;
		}
		break;
	    case 5:
		if (first.startsWith("#Data operacji")) {
		    state = 6;
		}
		break;
	    case 6:
		if (record.length > 6 && record[6].startsWith("#Saldo ko")) {
		    global.put("end_balance", record[4]);
		    state = 7;
		    continue;
		}
		if (record.length != 9)
		    continue;
		transactions.add(record);
		break;
	    }
	}
	return transactions;
    }

    static public String removeMultipleSpaces(String in) {
	String result = in;
	while (result.contains("  ")) {
	    result = result.replaceAll("  ", " ");
	}
	return result;
    }

    static public String removeOgonki(String in) {
	StringBuffer result = new StringBuffer();
	char[] arr = in.toCharArray();
	for (char a : arr) {
	    switch (a) {
	    case 'ą':
		a = 'a';
		break;
	    case 'ć':
		a = 'c';
		break;
	    case 'ę':
		a = 'e';
		break;
	    case 'ł':
		a = 'l';
		break;
	    case 'ń':
		a = 'n';
		break;
	    case 'ó':
		a = 'o';
		break;
	    case 'ś':
		a = 's';
		break;
	    case 'ź':
		a = 'z';
		break;
	    case 'ż':
		a = 'z';
		break;
	    case 'Ą':
		a = 'A';
		break;
	    case 'Ć':
		a = 'C';
		break;
	    case 'Ę':
		a = 'E';
		break;
	    case 'Ł':
		a = 'L';
		break;
	    case 'Ń':
		a = 'N';
		break;
	    case 'Ó':
		a = 'O';
		break;
	    case 'Ś':
		a = 'S';
		break;
	    case 'Ź':
		a = 'Z';
		break;
	    case 'Ż':
		a = 'Z';
		break;
	    }
	    result.append(a);
	}
	return result.toString();

    }
}
