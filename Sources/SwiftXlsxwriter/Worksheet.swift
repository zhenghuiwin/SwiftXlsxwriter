//
//  File.swift
//  
//
//  Created by zhenghuiwin on 2020/3/21.
//

import Foundation
import ExcelWriter

public class Worksheet {
    private var ws: UnsafeMutablePointer<lxw_worksheet>
    
    public init(_ ws: UnsafeMutablePointer<lxw_worksheet>) {
        self.ws = ws
    }
    
    //  worksheet_write_string(worksheet, row, UInt16(c), header[c], nil)
    public func writeString(row: Int,  col: Int, content: String, format: UnsafeMutablePointer<lxw_format>?) {
        worksheet_write_string(ws, UInt32(row), UInt16(col), content, format)
    }
}
