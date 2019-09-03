require "csv"
require "csv/excel/version"

class CSV
  module Excel
    UTF8BOM = -"\xEF\xBB\xBF"

    class Error < StandardError; end

    def initialize(data, **options)
      @for_excel = options.delete(:for_excel)
      options[:force_quotes] = true if @for_excel
      super(data, **options)
    end

    private
    def writer_options
      super.merge(for_excel: @for_excel)
    end
  end

  class Writer
    module Excel
      DATE_FORMAT = -"%Y/%m/%d"
      TIME_FORMAT = -"%Y/%m/%d %H:%M:%S"

      def quote_field(field)
        return super unless @options[:for_excel]

        case field
        when Date
          super(field.strftime(DATE_FORMAT))
        when Time, DateTime
          super(field.strftime(TIME_FORMAT))
        when String
          quoted = super
          encoded_assign_character = "=".encode(quoted.encoding)
          super(encoded_assign_character + quoted)
        else
          super
        end
      end
    end
  end
end

CSV.prepend CSV::Excel
CSV::Writer.prepend CSV::Writer::Excel
