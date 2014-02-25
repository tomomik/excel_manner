# encoding: utf-8

require "excel_manner/version"

require 'win32ole'
require 'yaml'

module ExcelManner
  extend self

  #
  # 処理開始
  #
  def start(yaml_path)

    @config = YAML.load_file(yaml_path)
    @app = WIN32OLE.new('Excel.Application')

    if @config["read_path"]

      print "処理開始!" + "\n"
      read_path  = File.expand_path(@config["read_path"])
      print "  変更したファイル\n"
      traverse_path(read_path)
      print "全てのファイルの処理終了!" + "\n"

      return true
    else
      print "対象のパスが設定されていません。" + "\n"
    end

  end

  private
  
  #
  # 引数 path 以下のファイルで
  # Excelファイルがあればmodify_excel_fileを実行する
  #
  def traverse_path(path)

    if FileTest.directory?(path)
      dir = Dir.open(path)
      while name = dir.read
        next if name == "."
        next if name == ".."
        traverse_path(path + "/" + name)
      end
      dir.close
    else
      if File.extname(path) == '.xls' or File.extname(path) == '.xlsx'
        modify_excel_file(path)
      end
    end

  end

  #
  # Excelの変更処理
  #
  def modify_excel_file(xls_path)

    begin

      print "\t" + File.basename(xls_path).to_s.encode("Shift_JIS") + "\n"

      book = @app.Workbooks.Open(xls_path)

      first_sheet_name = nil

      # 各シートの処理
      book.WorkSheets.each do |s|
        if first_sheet_name == nil then
          first_sheet_name = s.Name
        end

        s.select
        s.Range("A1").Select

        # 表示率
        @app.ActiveWindow.zoom = @config["zoom_ratio"]||100
        # 表示形式：標準
        @app.ActiveWindow.view = 1

      end

      # 最初のシートを選択
      first_sheet = book.WorkSheets.item(first_sheet_name)
      first_sheet.select

      # プロパティを変更
      # title
      book.builtinDocumentProperties("Title").value = ""
      # 表題
      book.builtinDocumentProperties("Subject").value = ""
      # 作成者
      book.builtinDocumentProperties("Author").value = ""
      # カテゴリ
      book.builtinDocumentProperties("Category").value = ""
      # キーワード
      book.builtinDocumentProperties("Keywords").value = ""
      # コメント
      book.builtinDocumentProperties("Comments").value = ""

      # 保存
      book.Save
      book.close
      @app.quit

    rescue => ex
      print "\n" + ex.message + "\n"
    ensure
      @app.quit
    end

  end

end
