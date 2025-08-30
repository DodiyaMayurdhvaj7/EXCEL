"use client";

import { useState, useMemo, ChangeEvent, useRef, useCallback } from "react";
import * as xlsx from "xlsx";
import { Upload, ListChecks, Download, FileSpreadsheet, X } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Separator } from "@/components/ui/separator";
import { useToast } from "@/hooks/use-toast";

type ExcelRow = { [key: string]: string | number };

export default function ExcelImporterExporterPage() {
  const [data, setData] = useState<ExcelRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [selectedField, setSelectedField] = useState<string>('');
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [fileName, setFileName] = useState<string>('');
  
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processData = useCallback((jsonData: ExcelRow[]) => {
    setData(jsonData);
    const fileHeaders = Object.keys(jsonData[0]);
    setHeaders(fileHeaders);
    setSelectedField(fileHeaders[0]);
    setSelectedRows(new Set());
  }, []);

  const handleFileUpload = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.name.match(/\.(xlsx|xls)$/)) {
      toast({
        variant: "destructive",
        title: "Invalid File Type",
        description: "Please upload a valid Excel file (.xlsx or .xls).",
      });
      return;
    }

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const fileData = e.target?.result;
        const workbook = xlsx.read(fileData, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: ExcelRow[] = xlsx.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
          toast({
            variant: "destructive",
            title: "Empty File",
            description: "The Excel file seems to be empty.",
          });
          return;
        }

        processData(jsonData);
        
        toast({
          title: "File Uploaded Successfully",
          description: `${jsonData.length} rows loaded from ${file.name}.`,
        });
      } catch (error) {
        console.error("Error parsing Excel file:", error);
        toast({
          variant: "destructive",
          title: "File Parsing Error",
          description: "Could not read or parse the Excel file. Please ensure it's not corrupted.",
        });
      }
    };
    reader.onerror = () => {
        toast({
          variant: "destructive",
          title: "File Read Error",
          description: "There was an error reading the file.",
        });
    }
    reader.readAsBinaryString(file);
    
    if (fileInputRef.current) {
        fileInputRef.current.value = "";
    }
  };

  const handleCheckboxChange = (index: number) => {
    const newSelectedRows = new Set(selectedRows);
    if (newSelectedRows.has(index)) {
      newSelectedRows.delete(index);
    } else {
      newSelectedRows.add(index);
    }
    setSelectedRows(newSelectedRows);
  };

  const handleExport = () => {
    if (selectedRows.size === 0) {
      toast({
        variant: "destructive",
        title: "No Products Selected",
        description: "Please select at least one product to export.",
      });
      return;
    }

    try {
      const selectedData = Array.from(selectedRows).map(index => data[index]);
      
      const worksheet = xlsx.utils.json_to_sheet(selectedData);
      const workbook = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(workbook, worksheet, 'Selected Products');
      xlsx.writeFile(workbook, 'selected_products.xlsx');

      toast({
        title: "Success!",
        description: `Exported ${selectedRows.size} products to selected_products.xlsx.`,
      });

    } catch (error) {
      console.error("Error during export:", error);
      toast({
        variant: "destructive",
        title: "An Error Occurred",
        description: "Could not export file. Please try again.",
      });
    }
  };

  const handleReset = () => {
    setData([]);
    setHeaders([]);
    setSelectedField('');
    setSelectedRows(new Set());
    setFileName('');
    if (fileInputRef.current) {
        fileInputRef.current.value = "";
    }
  };

  const visibleProducts = useMemo(() => {
    if (!selectedField) return [];
    return data.map((row, index) => ({
      id: index,
      label: row[selectedField] || 'N/A',
    }));
  }, [data, selectedField]);

  return (
    <main className="min-h-screen w-full">
      <div className="container mx-auto p-4 sm:p-6 md:p-10">
        <header className="text-center mb-10">
          <h1 className="text-4xl lg:text-5xl font-headline font-bold tracking-tight">
            Excel Importer & Exporter
          </h1>
          <p className="text-lg text-muted-foreground mt-2 max-w-2xl mx-auto">
            Upload, select, and export your product data.
          </p>
        </header>

        {data.length === 0 ? (
          <Card className="max-w-xl mx-auto border-2 border-dashed hover:border-primary transition-colors duration-300">
            <CardContent className="p-10 text-center">
              <input
                type="file"
                ref={fileInputRef}
                className="hidden"
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
              />
              <FileSpreadsheet className="mx-auto h-16 w-16 text-muted-foreground mb-4" />
              <h3 className="text-2xl font-semibold mb-2">Upload Your Excel File</h3>
              <p className="text-muted-foreground mb-6">Supports .xlsx and .xls formats. Drag & drop or click to select.</p>
              <Button size="lg" onClick={() => fileInputRef.current?.click()}>
                <Upload className="mr-2 h-5 w-5" />
                Select File
              </Button>
            </CardContent>
          </Card>
        ) : (
          <div className="flex flex-col gap-8">
            <Card>
                <CardHeader className="flex flex-row items-center justify-between">
                    <div>
                        <CardTitle className="flex items-center gap-2">
                           <FileSpreadsheet className="h-5 w-5"/>
                           {fileName}
                        </CardTitle>
                        <CardDescription>
                            {data.length} rows loaded. Select products below.
                        </CardDescription>
                    </div>
                    <Button variant="outline" size="sm" onClick={handleReset}><X className="mr-2 h-4 w-4"/>Clear File</Button>
                </CardHeader>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <ListChecks />
                  Select Products
                </CardTitle>
                <CardDescription>
                  {selectedRows.size} of {data.length} products selected.
                </CardDescription>
              </CardHeader>
              <CardContent className="flex flex-col gap-4">
                  <div className="flex items-center gap-4">
                      <Label htmlFor="field-select" className="whitespace-nowrap">Display Field:</Label>
                      <Select value={selectedField} onValueChange={setSelectedField}>
                      <SelectTrigger id="field-select" className="w-full">
                          <SelectValue placeholder="Select a field" />
                      </SelectTrigger>
                      <SelectContent>
                          {headers.map((header) => (
                          <SelectItem key={header} value={header}>{header}</SelectItem>
                          ))}
                      </SelectContent>
                      </Select>
                  </div>
                <Separator />
                <ScrollArea className="h-96 border rounded-md p-4">
                  <div className="space-y-4">
                    {visibleProducts.map((product) => (
                      <div key={product.id} className="flex items-center space-x-3 p-2 rounded-md hover:bg-secondary">
                        <Checkbox
                          id={`product-${product.id}`}
                          checked={selectedRows.has(product.id)}
                          onCheckedChange={() => handleCheckboxChange(product.id)}
                        />
                        <Label htmlFor={`product-${product.id}`} className="flex-1 cursor-pointer font-normal text-sm">
                          {product.label}
                        </Label>
                      </div>
                    ))}
                  </div>
                </ScrollArea>
              </CardContent>
              <CardFooter>
                  <Button
                    size="lg"
                    className="w-full bg-accent hover:bg-accent/90 text-accent-foreground"
                    onClick={handleExport}
                    disabled={selectedRows.size === 0}
                  >
                    <Download className="mr-2 h-5 w-5" />
                    Export Selected ({selectedRows.size})
                  </Button>
                </CardFooter>
            </Card>
          </div>
        )}
      </div>
    </main>
  );
}
