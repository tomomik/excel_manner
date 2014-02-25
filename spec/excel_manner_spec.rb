# coding: utf-8

require './spec_helper.rb'

describe ExcelManner do

  context '正常系' do
    it "正常終了" do
      expect(ExcelManner.start("./data/config001.yml")).to eq(true)
    end

  end

end
