require "spec_helper"

describe CSV::Excel do
  let(:out) { "" }
  let(:csv) {
    CSV.new(out, for_excel: true)
  }
  let(:row) {
    ["012", "12", "abc", "foo\"\nbar", 123, Date.new(2019, 4, 1), Time.new(2019, 4, 1, 12, 34, 56)]
  }

  subject {
    csv << row
    out
  }

  it "use double quoted formula for string value" do
    should =~ %r(\A"=""012""","=""12""","=""abc""","=""foo""""\nbar""",)
  end

  it "don't use formula for string value" do
    should =~ %r(,"123",)
  end

  it "use yyyy/mm/dd format for date value" do
    should =~ %r(,"2019/04/01",)
  end

  it "use yyyy/mm/dd hh:mm:ss format for date value" do
    should =~ %r(,"2019/04/01 12:34:56"$)
  end
end
