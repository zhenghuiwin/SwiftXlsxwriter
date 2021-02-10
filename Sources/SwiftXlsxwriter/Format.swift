//
//  File.swift
//  
//
//  Created by zhenghuiwin on 2020/3/21.
//

import Foundation
import ExcelWriter

public struct Format {
    private let fmt: UnsafeMutablePointer<lxw_format>
    
    var rawValue: UnsafeMutablePointer<lxw_format> {
        get {
            return fmt
        }
    }
    
    public init(_ fmt: UnsafeMutablePointer<lxw_format>) {
        self.fmt = fmt
    }
    
    public func setBold() {
        format_set_bold(fmt)
    }
    
    public func setBorder(type: Borders) {
        format_set_border(fmt, type.rawValue)
    }
    
    public func setAlign(type: Alignments) {
        format_set_align(fmt, type.rawValue)
    }
    
    public func setFont(size: Double) {
        format_set_font_size(fmt, size)
    }
    
    public func set(bgColor: UInt32) {
        format_set_bg_color(fmt, bgColor)
    }
    
    
}

public enum Borders: UInt8 {
    case none = 0
    case thin = 1
    case medium = 2
    case dashed = 3
    case dotted = 4
    case thick = 5
    case double = 6
}

public enum Alignments: UInt8 {
    case none = 0
    case left = 1
    case center = 2
    case right = 3
    case fill = 4
    case justify = 5
    case centerAcross = 6
    case distributed = 7
    case verticalTop = 8
    case verticalBottom = 9
    case verticalCenter = 10
    case verticalJustify = 11
    case verticalDistributed = 12
}
