require 'rubygems'
require 'serialport'

class ProXR
  attr_accessor :serial_port

  SUCCESS = 85
  def initialize(device = "/dev/ttyS0", baud_rate = 115200)
    @serial_port = SerialPort.new device, baud_rate
  end

  def send_command(*cmds)
    serial_port.write 254.chr
    cmds.each {|cmd| serial_port.write cmd.chr }

    serial_port.getc
  end

  def reporting_mode
    send_command(27)
  end

  def getc
    serial_port.getc
  end

  def read_voltage(relay_number, bank_number)
    send_command(150, bank_number)
  end

  def relay_on(relay_number, bank_number)
    relay_on_cmd = (108+relay_number)
    send_command relay_on_cmd, bank_number
  end

  def relay_off(relay_number, bank_number)
    relay_off_cmd = (100+relay_number)
    send_command relay_off_cmd, bank_number
  end




end

if $0 == __FILE__
  require 'test/unit'
  class TestReportingMode < Test::Unit::TestCase
    def setup
      @serial_port = ProXR.new
    end
    def test_should_indicate_it_is_in_reporting_mode
      assert_equal ProXR::SUCCESS, @serial_port.reporting_mode
    end

    def test_voltage_at_0_1_should_be_zero
      assert_equal 0, @serial_port.read_voltage(0, 0)
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

