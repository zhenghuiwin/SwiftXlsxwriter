//
//  File.swift
//  
//
//  Created by zhenghuiwin on 2020/3/21.
//

import Foundation
import ExcelWriter

public class Workbook {
    private var wb: UnsafeMutablePointer<lxw_workbook>
    
    public init(file: String) {
        wb = workbook_new(file)
    }
    
    deinit {
        workbook_close(wb)
    }
    
    // workbook_add_worksheet
    public func addWorksheet(sheetName: String?) -> Worksheet? {
        guard let sheet = workbook_add_worksheet(wb, sheetName) else {
            return nil
        }
        
        return  Worksheet(sheet)
    }
}
