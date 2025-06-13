import express, { Request, Response } from 'express';
import { format, parse, isValid, isEqual } from 'date-fns';
import Joi from 'joi';
import * as XLSX from 'xlsx';
import multer from 'multer';

const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());

const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (_req: Express.Request, file: Express.Multer.File, cb: multer.FileFilterCallback) => {
    const allowedMimeTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];

    if (allowedMimeTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files are allowed'));
    }
  }
});

enum InvoiceParsingStage {
    INVOICING_MONTH,
    CURRENCY_RATES,
    INVOICES_DATA_COLUMNS,
    INVOICES_DATA_ROWS,
}

const uploadSchema = Joi.object({
  invoicingMonth: Joi.string()
    .pattern(/^\d{4}-(0[1-9]|1[0-2])$/)
    .required()
    .messages({
      'string.pattern.base': 'invoicingMonth must be in YYYY-MM format',
      'any.required': 'invoicingMonth is required'
    })
});

const invoiceRowSchema = Joi.object({
    Customer: Joi.string().required().messages({
        'any.required': 'Customer is required',
        'string.base': 'Customer must be a string',
    }),
    "Cust No'": Joi.alternatives().try(Joi.string(), Joi.number()).required().messages({
        'any.required': 'Cust No is required',
    }),
    'Project Type': Joi.string().required().messages({
        'any.required': 'Project Type is required',
    }),
    Quantity: Joi.number().required().messages({
        'any.required': 'Quantity is required',
        'number.base': 'Quantity must be a number',
    }),
    'Price Per Item': Joi.number().required().messages({
        'any.required': 'Price Per Item is required',
        'number.base': 'Price Per Item must be a number',
    }),
    'Item Price Currency': Joi.string().required().messages({
        'any.required': 'Item Price Currency is required',
    }),
    'Invoice Total Price': Joi.number().required().messages({
        'any.required': 'Invoice Total Price is required',
        'number.base': 'Invoice Total Price must be a number',
    }),
    'Invoice Currency': Joi.string().required().messages({
        'any.required': 'Invoice Currency is required',
    }),
    Status: Joi.string().required().messages({
        'any.required': 'Status is required',
    }),
}).unknown(true);

app.post('/upload', upload.single('file'), (req: Request & { file?: Express.Multer.File }, res: Response) => {    
    try {
        const { invoicingMonth, file } = getParsedBody(req);

        const workbook = XLSX.read(file.buffer, { type: 'buffer' });

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 });

        let currentStage = InvoiceParsingStage.INVOICING_MONTH;
        let actualInvoicingMonth: Date;

        let dataColumns: string[] = [];

        const currencyRates = new Map<string, number>();
        const invoicesData = [];

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            if (!row.length) {
                if (currentStage === InvoiceParsingStage.INVOICING_MONTH || currentStage === InvoiceParsingStage.CURRENCY_RATES || currentStage === InvoiceParsingStage.INVOICES_DATA_COLUMNS) {
                    throw new Error('Unexpected empty row for stage ' + currentStage);
                }

                break;
            }

            if (currentStage === InvoiceParsingStage.INVOICING_MONTH) {
                if (row.length === 1 && typeof row[0] === 'string') {
                    actualInvoicingMonth = parseInvoiceMonth(row[0]);

                    if (!isEqual(invoicingMonth, actualInvoicingMonth)) {
                        throw new Error('Passed invoicing month doesn\'t match the actual invoicing month')
                    }

                    currentStage = InvoiceParsingStage.CURRENCY_RATES;

                    continue;
                }

                throw new Error('Unexpected row for stage ' + currentStage);
            }

            if (currentStage === InvoiceParsingStage.CURRENCY_RATES) {
                if (row.length === 2 && typeof row[0] === 'string' && typeof row[1] === 'number') {
                    const symbol = parseCurrencySymbol(row[0]);

                    if (currencyRates.has(symbol)) {
                        throw new Error('Duplicate currency symbol: ' + symbol);
                    }

                    currencyRates.set(symbol, row[1]);

                    continue;
                } else if (!currencyRates.size) {
                    throw new Error('Unexpected row for stage ' + currentStage);
                } else {
                    currentStage = InvoiceParsingStage.INVOICES_DATA_COLUMNS;
                    i--;

                    continue;
                }
            }

            if (currentStage === InvoiceParsingStage.INVOICES_DATA_COLUMNS) {
                if (!row.every((cell) => typeof cell === 'string')) {
                    throw new Error('Unexpected row for stage ' + currentStage);
                }

                dataColumns = row.map((cell) => cell.trim());
                currentStage = InvoiceParsingStage.INVOICES_DATA_ROWS;

                continue;
            }


            if (currentStage === InvoiceParsingStage.INVOICES_DATA_ROWS) {
                if (!row.length) {
                    if (!invoicesData.length) {
                        throw new Error('Unexpected empty row for stage ' + currentStage);
                    }

                    break;
                }

                const rawItem: Record<string, any> = {
                    'Invoice Total': null,
                    validationErrors: [],
                };

                dataColumns.forEach((column, index) => {
                    rawItem[column] = row[index];
                });

                const rowValidationErrors = invoiceRowSchema.validate(rawItem, { abortEarly: false });

                if (rowValidationErrors.error) {
                    rawItem.validationErrors.push(
                        ...rowValidationErrors.error.details.map((detail: Joi.ValidationErrorItem) => detail.message)
                    );
                }

                const isValidItem = rawItem['Status'] === 'Ready' || (rawItem['Invoice #'] as string)?.length;

                if (!isValidItem) {
                    continue;
                }

                const invoicesTotalResult = getInvoiceTotal(rawItem['Invoice Total Price'], rawItem['Invoice Currency'], currencyRates);

                if (invoicesTotalResult.errors.length) {
                    rawItem.validationErrors.push(...invoicesTotalResult.errors);
                } else {
                    rawItem['Invoice Total'] = invoicesTotalResult.value;
                }

                invoicesData.push(rawItem);
            }
        }

        console.log(actualInvoicingMonth!);
        res.json({
            //   check time zone
            invoicingMonth: format(actualInvoicingMonth!, 'yyyy-MM'),
            currencyRates: Object.fromEntries(currencyRates),
            invoicesData,
        });

  } catch (error) {
    console.error('Error processing Excel file:', error);

    // inaccurate status code, actually there are cases for 400s and 500s
    res.status(422).json({ error: error instanceof Error ? error.message : 'Unknown error' });
  }
});

const getParsedBody = (req: Request & { file?: Express.Multer.File }) => {
    const { error, value } = uploadSchema.validate(req.body);

    if (error) {
        throw new Error(error.details[0].message);
    }

    if (!req.file) {
        throw new Error('No file uploaded');
    }

    const invoicingMonth = parse(value.invoicingMonth, 'yyyy-MM', new Date());

    if (!isValid(invoicingMonth)) {
        throw new Error('Invalid invoicing month');
    }

    return {
        invoicingMonth,
        file: req.file,
    };
}   

const parseInvoiceMonth = (value: string): Date => {
    const formats = ['MM yyyy', 'MMM yyyy', 'M yyyy'];

    for (const format of formats) {
        const parsed = parse(value, format, new Date());

        if (isValid(parsed)) {
            return parsed;
        }
    }

    throw new Error('Unable to parse invoice date');
}

const parseCurrencySymbol = (value: string): string => {
    if (value.match(/^[A-Z]{3}$/)) {
        return value;
    }

    if (value.includes('Rate')) {
        const symbol = value.split('Rate')[0].trim();

        if (symbol.length) {
            return symbol;
        }
    }

    throw new Error('Unable to parse currency symbol');
}

const getInvoiceTotal = (price: any, currency: string | undefined, currencyRates: Map<string, number>): { errors: string[], value: number | null } => {
    const errors: string[] = [];
    let total = null;
    
    if (currency && !currencyRates.has(currency)) {
        errors.push('Currency rate not found');
    }

    if (!errors.length && typeof currency === 'string' && typeof price === 'number') {
        total = price * currencyRates.get(currency)!;
    }

    return { errors, value: total };
}

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
}); 