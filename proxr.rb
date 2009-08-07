require 'rubygems'
require 'serialport'
require 'timeout'

class ProXR < SerialPort
  SUCCESS = 85

  def initialize(device = "/dev/ttyS0", *params)
    default_params = {:baud_rate => 115200}
    super device, default_params.merge(params)
  end

  def send_command(*cmds)
    Timeout::timeout(0.5) do
      write 254.chr
      cmds.each {|cmd| write cmd.chr }

      getc
    end
  end

  def reporting_mode
    send_command(27)
  end

  def read_voltage(channel)
    send_command(150 + channel)
  end

  def relay_on(relay_number, bank_number)
    relay_on_cmd = (108+relay_number)
    send_command relay_on_cmd, bank_number
  end

  def relay_off(relay_number, bank_number)
    relay_off_cmd = (100+relay_number)
    send_command relay_off_cmd, bank_number
  end

  def show_all_voltages
    p = ProXR.new
    puts "Ch:Volt"
    (0..7).each do |channel|
      puts " #{channel}:#{p.read_voltage(channel)}"
    end
  end

end

if $0 == __FILE__
  require 'test/unit'
  class TestReportingMode < Test::Unit::TestCase
    def setup
      @serial_port = ProXR.new "/dev/ttyS0"
    end

    def test_should_indicate_it_is_in_reporting_mode
      assert_equal ProXR::SUCCESS, @serial_port.reporting_mode
    end

    def test_voltage_at_0_1_should_be_255
      assert_equal 255, @serial_port.read_voltage(0)
    end

    def test_relay_on_0_1
      assert_equal ProXR::SUCCESS, @serial_port.relay_on(0, 1)
    end

    def test_relay_off_0_1
      assert_equal ProXR::SUCCESS, @serial_port.relay_off(0, 1)
    end

#    def set_relay_status(:red, :green)
#
#    end
  end
end

