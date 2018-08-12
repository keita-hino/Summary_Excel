class ConfigReader
  def initialize
    @date = [{}]
    @date.delete_at(0)
  end
  def add(id,name)
     @date.push(
      "id":id,
      "name":name
    )
  end
  def read
    @date
  end
end
