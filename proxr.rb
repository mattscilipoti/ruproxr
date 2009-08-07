require 'rubygems'
require 'serialport'
require 'timeout'

class ProXR < SerialPort
  SUCCESS = 85

  #SerialPort uses ::new instead of initialize.
  def ProXR::new(port = "/dev/ttyS0", *params)
    default_params = [115200]
    default_params.each_with_index {|param, index| params[index] = param unless params[index] }
    super port, *params
#    read_timeout = 2 #does this do anything?
  end

  def has_voltage?(channel)
    read_voltage(channel) > 0 ? true : false
  end

  def send_command(*cmds)
    Timeout::timeout(1) do
      write 254.chr
      cmds.each {|cmd| write cmd.chr }
      getc
    end
  end

  def reporting_mode
    send_command(27)
  end

  def read_voltage(channel, data_bits = 8)
    case data_bits
      when 8
      max_voltage = 5
      max_value = max_reading_for_8_bit = 255
      cmd = 150
    end

    voltage_conversion_factor = (max_value/max_voltage)
    ad_voltage = send_command(cmd + channel)
    voltage = ad_voltage / voltage_conversion_factor

    #special case for 8-bit
    voltage = 0 if data_bits == 8 && ad_voltage == 255
    voltage
end

  def relay_on(relay_number, bank_number)
    relay_on_cmd = (108+relay_number)
    send_command relay_on_cmd, bank_number
  end

  def relay_on?(relay_number, bank_number)
    relay_status(relay_number, bank_number) == 1
  end

  def relay_off(relay_number, bank_number)
    relay_off_cmd = (100+relay_number)
    send_command relay_off_cmd, bank_number
  end

  def relay_status(relay_number, bank_number)
    relay_status_cmd = (116+relay_number)
    send_command relay_status_cmd, bank_number
  end

  def show_all_voltages
#    puts "Ch:Volt"
    (0..7).collect do |channel|
      "#{channel}:#{read_voltage(channel)}"
    end
  end

end

if $0 == __FILE__
  require 'test/unit'
  class TestProXR < Test::Unit::TestCase
    def setup
      @serial_port = ProXR.new
    end

    def test_should_indicate_it_is_in_reporting_mode
      assert_equal ProXR::SUCCESS, @serial_port.reporting_mode
    end

    def test_voltage_at_0_1_should_be_zero
      assert_equal 0, @serial_port.read_voltage(0)
    end

    def test_relay_on_0_1
      assert_equal ProXR::SUCCESS, @serial_port.relay_on(0, 1)
    end

    def test_relay_off_0_1
      assert_equal ProXR::SUCCESS, @serial_port.relay_off(0, 1)
    end

    def test_has_voltage_without_voltage
      @serial_port.instance_eval do
        def read_voltage(*args)
          0
        end
      end

      assert !@serial_port.has_voltage?(0)
    end

    def test_has_voltage_wit_voltage
      @serial_port.instance_eval do
        def read_voltage(*args)
          2
        end
      end

      assert @serial_port.has_voltage?(0)
    end

    def test_relay_status_for_on
      @serial_port.relay_on(0, 1)
      assert_equal 1, @serial_port.relay_status(0, 1)
    end

    def test_relay_status_for_off
      @serial_port.relay_off(0, 1)
      assert_equal 0, @serial_port.relay_status(0, 1)
    end

    def test_when_on_relay_on?
      @serial_port.relay_on(0, 1)
      assert @serial_port.relay_on?(0, 1)
    end

    def test_when_off_relay_on?
      @serial_port.relay_off(0, 1)
      assert !@serial_port.relay_on?(0, 1)
    end
  end
end

